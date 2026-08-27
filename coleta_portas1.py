"""
Coleta o status das portas fisicas de switches Cisco, Fortinet/FortiSwitch e Huawei
via SSH/Telnet (Netmiko), em paralelo, e exporta os resultados em CSV e JSON.

Versao ajustada para Fortinet/FortiSwitch no formato:
== [ port1 ]
name: port1    egress-drop-mode: enabled    link-status: up (1000Mbps full-duplex)   status: up

Uso basico:
    python coleta_portas_fortinet_final.py

Uso com parametros:
    python coleta_portas_fortinet_final.py -i switches.csv -o resultado_portas_fisicas.csv -j dados_portas.json -w 5

Inventario esperado (switches.csv):
    name,host,device_type,username,password,secret
    SW-CISCO-SSH-01,10.10.10.1,cisco_ios,admin,senha,enable
    SW-CISCO-TELNET-01,10.10.10.2,cisco_ios_telnet,admin,senha,enable
    FSW-01,10.10.10.3,fortinet,admin,senha,
    HUA-SW-01,10.10.10.4,huawei,admin,senha,
    HUA-SW-02,10.10.10.5,huawei_telnet,admin,senha,
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Optional

from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoAuthenticationException, NetmikoTimeoutException


# =========================================================
# CONFIGURACAO / LOGGING
# =========================================================

LOG = logging.getLogger("coleta_portas")

DEVICE_TYPES_TELNET = {"cisco_ios_telnet", "huawei_telnet"}
DEVICE_TYPES_CISCO = {"cisco_ios", "cisco_ios_telnet"}
DEVICE_TYPES_HUAWEI = {"huawei", "huawei_telnet", "huawei_vrpv8"}
DEVICE_TYPES_FORTINET = {"fortinet"}
DEVICE_TYPES_SUPORTADOS = DEVICE_TYPES_CISCO | DEVICE_TYPES_HUAWEI | DEVICE_TYPES_FORTINET

CSV_FIELDS = [
    "data_hora",
    "switch",
    "host",
    "device_type",
    "porta",
    "status",
    "raw_status",
]


def configurar_logging(verbose: bool = False) -> None:
    nivel = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=nivel,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler("coleta_portas.log", encoding="utf-8"),
        ],
    )


# =========================================================
# MODELOS
# =========================================================

@dataclass
class SwitchConfig:
    name: str
    host: str
    device_type: str
    username: str
    password: str
    secret: str = ""


@dataclass
class PortaResultado:
    data_hora: str
    switch: str
    host: str
    device_type: str
    porta: str
    status: str
    raw_status: str


# =========================================================
# INVENTARIO
# =========================================================

def carregar_switches(arquivo: str) -> list[SwitchConfig]:
    caminho = Path(arquivo)

    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo de inventario nao encontrado: {arquivo}")

    switches: list[SwitchConfig] = []

    with caminho.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)

        campos_obrigatorios = {"name", "host", "device_type", "username", "password"}
        campos_arquivo = set(reader.fieldnames or [])
        faltando = campos_obrigatorios - campos_arquivo

        if faltando:
            raise ValueError(
                "Inventario invalido. Campos obrigatorios ausentes: "
                + ", ".join(sorted(faltando))
            )

        for linha_num, row in enumerate(reader, start=2):
            name = (row.get("name") or "").strip()
            host = (row.get("host") or "").strip()
            device_type = (row.get("device_type") or "").strip()
            username = (row.get("username") or "").strip()
            password = (row.get("password") or "").strip()
            secret = (row.get("secret") or "").strip()

            if not name or not host or not device_type or not username or not password:
                LOG.warning("Linha %s ignorada: campos obrigatorios vazios", linha_num)
                continue

            if device_type not in DEVICE_TYPES_SUPORTADOS:
                LOG.warning(
                    "Linha %s ignorada: device_type nao suportado: %s",
                    linha_num,
                    device_type,
                )
                continue

            switches.append(
                SwitchConfig(
                    name=name,
                    host=host,
                    device_type=device_type,
                    username=username,
                    password=password,
                    secret=secret,
                )
            )

    if not switches:
        raise ValueError("Nenhum switch valido encontrado no inventario.")

    return switches


# =========================================================
# FILTROS - PORTAS FISICAS
# =========================================================

def porta_fisica_cisco(porta: str) -> bool:
    porta = porta.strip()

    prefixos_ignorar = (
        "Vl", "Vlan", "Lo", "Loopback", "Nu", "Null", "Po",
        "Port-channel", "Tu", "Tunnel", "Mgmt", "Management",
    )

    prefixos_fisicos = (
        "Fa", "FastEthernet",
        "Gi", "GigabitEthernet",
        "Te", "TenGigabitEthernet",
        "Eth", "Ethernet",
        "Fo", "FortyGigabitEthernet",
        "Hu", "HundredGigE",
        "Twe", "TwentyFiveGigE",
        "Tw",
        "GE", "XGE", "10GE", "25GE", "40GE", "100GE",
    )

    if porta.startswith(prefixos_ignorar):
        return False

    return porta.startswith(prefixos_fisicos)


def porta_fisica_fortinet(porta: str) -> bool:
    porta = porta.strip().lower()

    prefixos_ignorar = (
        "internal",
        "mgmt",
        "ha",
        "fortilink",
        "trunk",
        "vlan",
        "npu",
    )

    if porta.startswith(prefixos_ignorar):
        return False

    padroes_fisicos = (
        r"^port\d+$",
        r"^sfp\d+$",
        r"^x\d+$",
        r"^ge\d+$",
        r"^10ge\d+$",
        r"^25ge\d+$",
        r"^40ge\d+$",
        r"^100ge\d+$",
    )

    return any(re.match(padrao, porta) for padrao in padroes_fisicos)


def porta_fisica_huawei(porta: str) -> bool:
    porta = porta.strip()

    prefixos_ignorar = (
        "Vlanif", "LoopBack", "NULL", "Eth-Trunk", "MEth", "Tunnel",
    )

    prefixos_fisicos = (
        "GigabitEthernet", "XGigabitEthernet", "GE", "XGE",
        "10GE", "25GE", "40GE", "100GE", "Ethernet",
    )

    if porta.startswith(prefixos_ignorar):
        return False

    return porta.startswith(prefixos_fisicos)


# =========================================================
# CISCO
# =========================================================

def normalizar_status_cisco(status: str) -> str:
    status = status.strip().lower()

    if status == "connected":
        return "UP"

    if status in {"notconnect", "disabled", "inactive", "err-disabled"}:
        return "DOWN"

    return status.upper()


def parse_cisco_show_interfaces_status(output: str) -> list[dict]:
    resultados: list[dict] = []
    status_possiveis = ("connected", "notconnect", "disabled", "inactive", "err-disabled")

    for linha in output.splitlines():
        linha = linha.rstrip()

        if not linha:
            continue

        if linha.startswith("Port") or linha.startswith("--"):
            continue

        linha_lower = linha.lower()

        if not any(st in linha_lower for st in status_possiveis):
            continue

        partes = re.split(r"\s+", linha.strip())

        if len(partes) >= 2:
            porta = partes[0]

            if not porta_fisica_cisco(porta):
                continue

            status_encontrado = None
            for item in partes:
                item_lower = item.lower()
                if item_lower in status_possiveis:
                    status_encontrado = item_lower
                    break

            if not status_encontrado:
                continue

            status = normalizar_status_cisco(status_encontrado)

            resultados.append(
                {
                    "porta": porta,
                    "status": status,
                    "raw_status": status_encontrado,
                }
            )

    return resultados


# =========================================================
# FORTINET / FORTISWITCH
# =========================================================

def normalizar_status_fortinet(link_status: str) -> str:
    if not link_status:
        return "UNKNOWN"

    texto = link_status.strip().lower()

    # link-status real da porta: up/down
    if texto.startswith("up"):
        return "UP"

    if texto.startswith("down"):
        return "DOWN"

    if "up" in texto:
        return "UP"

    if "down" in texto:
        return "DOWN"

    return texto.upper()


def parse_fortinet_ports(output: str) -> list[dict]:
    """
    Parser Fortinet/FortiSwitch para saida no formato real:

    == [ port1 ]
    name: port1    egress-drop-mode: enabled    link-status: up (1000Mbps full-duplex)   status: up
    == [ port2 ]
    name: port2    egress-drop-mode: enabled    link-status: down   status: up

    Importante:
    - Usa link-status como prioridade.
    - Ignora o campo status quando link-status existe, pois status indica admin/up e nao link fisico.
    """
    resultados: list[dict] = []
    porta_atual: Optional[str] = None

    for linha in output.splitlines():
        linha = linha.strip()

        if not linha:
            continue

        # Exemplo: == [ port1 ]
        m_bloco = re.match(r"==\s*\[\s*(.+?)\s*\]", linha)
        if m_bloco:
            porta_atual = m_bloco.group(1).strip()
            continue

        # Exemplo:
        # name: port1    egress-drop-mode: enabled    link-status: up (1000Mbps full-duplex)   status: up
        if "name:" in linha.lower():
            nome_match = re.search(r"\bname:\s*(\S+)", linha, re.IGNORECASE)

            if nome_match:
                porta = nome_match.group(1).strip()
            else:
                porta = porta_atual

            if not porta:
                continue

            if not porta_fisica_fortinet(porta):
                continue

            # Captura link-status ate antes de "status:" ou fim da linha.
            # Ex: "up (1000Mbps full-duplex)" ou "down"
            link_match = re.search(
                r"\blink-status:\s*(.*?)(?:\s+status:|$)",
                linha,
                re.IGNORECASE,
            )

            status_match = re.search(
                r"\bstatus:\s*(\S+)",
                linha,
                re.IGNORECASE,
            )

            if link_match:
                raw_status = link_match.group(1).strip()
            elif status_match:
                raw_status = status_match.group(1).strip()
            else:
                raw_status = "UNKNOWN"

            status = normalizar_status_fortinet(raw_status)

            resultados.append(
                {
                    "porta": porta,
                    "status": status,
                    "raw_status": raw_status,
                }
            )

    return resultados


# =========================================================
# HUAWEI
# =========================================================

def normalizar_status_huawei(phy: str, protocol: str) -> str:
    phy = phy.strip().lower()
    protocol = protocol.strip().lower()

    if phy == "up" and protocol.startswith("up"):
        return "UP"

    if phy.startswith("*down"):
        return "ADMIN_DOWN"

    if phy.startswith("down") or protocol.startswith("down"):
        return "DOWN"

    return f"{phy}/{protocol}".upper()


def parse_huawei_display_interface_brief(output: str) -> list[dict]:
    resultados: list[dict] = []

    prefixos_ignorados_linha = (
        "PHY:", "*down:", "^down:", "(l):", "(s):", "(e):", "(d):",
        "InUti/OutUti:", "Interface",
    )

    for linha in output.splitlines():
        linha = linha.strip()

        if not linha:
            continue

        if linha.startswith(prefixos_ignorados_linha):
            continue

        partes = re.split(r"\s+", linha)

        if len(partes) >= 3:
            porta = partes[0]

            if not porta_fisica_huawei(porta):
                continue

            phy = partes[1]
            protocol = partes[2]
            status = normalizar_status_huawei(phy, protocol)

            resultados.append(
                {
                    "porta": porta,
                    "status": status,
                    "raw_status": f"PHY={phy} PROTOCOL={protocol}",
                }
            )

    return resultados


# =========================================================
# COLETA
# =========================================================

def montar_device_dict(sw: SwitchConfig, timeout: int) -> dict:
    device = {
        "device_type": sw.device_type,
        "host": sw.host,
        "username": sw.username,
        "password": sw.password,
        "fast_cli": False,
        "conn_timeout": timeout,
        "banner_timeout": timeout + 5,
        "timeout": timeout * 3,
    }

    if sw.secret:
        device["secret"] = sw.secret

    if sw.device_type in DEVICE_TYPES_TELNET:
        device["port"] = 23

    return device


def executar_comando_por_tipo(conn, sw: SwitchConfig) -> list[dict]:
    if sw.device_type in DEVICE_TYPES_CISCO:
        output = conn.send_command("show interfaces status", read_timeout=60)
        return parse_cisco_show_interfaces_status(output)

    if sw.device_type in DEVICE_TYPES_FORTINET:
        output = conn.send_command(
            "get switch physical-port",
            expect_string=r"[#$]",
            read_timeout=60,
        )

        with open(f"debug_fortinet_{sw.name}_{sw.host}.txt", "w", encoding="utf-8") as f:
            f.write(output)

        return parse_fortinet_ports(output)

    if sw.device_type in DEVICE_TYPES_HUAWEI:
        output = conn.send_command("display interface brief", read_timeout=60)
        return parse_huawei_display_interface_brief(output)

    raise ValueError(f"device_type nao suportado: {sw.device_type}")


def coletar_switch(
    sw: SwitchConfig,
    timeout: int = 20,
    tentativas: int = 3,
    espera_retry: float = 3.0,
) -> list[PortaResultado]:
    """Conecta em um switch, coleta e normaliza o status das portas fisicas."""

    ultima_excecao: Optional[Exception] = None

    for tentativa in range(1, tentativas + 1):
        conn = None

        try:
            LOG.debug("Conectando em %s (%s), tentativa %s", sw.name, sw.host, tentativa)

            device = montar_device_dict(sw, timeout)
            conn = ConnectHandler(**device)

            if sw.device_type in DEVICE_TYPES_CISCO and sw.secret:
                conn.enable()

            portas = executar_comando_por_tipo(conn, sw)
            agora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            resultados = [
                PortaResultado(
                    data_hora=agora,
                    switch=sw.name,
                    host=sw.host,
                    device_type=sw.device_type,
                    porta=p["porta"],
                    status=p["status"],
                    raw_status=p["raw_status"],
                )
                for p in portas
            ]

            LOG.info("[OK] %s (%s) - %s portas fisicas", sw.name, sw.host, len(resultados))
            return resultados

        except NetmikoAuthenticationException as e:
            LOG.error("[ERRO] %s (%s) - autenticacao falhou: %s", sw.name, sw.host, e)
            raise

        except NetmikoTimeoutException as e:
            ultima_excecao = e
            LOG.warning(
                "[TIMEOUT] %s (%s) - tentativa %s/%s: %s",
                sw.name,
                sw.host,
                tentativa,
                tentativas,
                e,
            )

            if tentativa < tentativas:
                time.sleep(espera_retry)

        except Exception as e:
            ultima_excecao = e
            LOG.error("[ERRO] %s (%s) - %s", sw.name, sw.host, e)
            raise

        finally:
            if conn is not None:
                try:
                    conn.disconnect()
                except Exception:
                    pass

    raise RuntimeError(f"Falha apos {tentativas} tentativa(s): {ultima_excecao}")


# =========================================================
# SALVAR RESULTADOS
# =========================================================

def salvar_csv(resultados: list[PortaResultado], arquivo_saida: str) -> None:
    with open(arquivo_saida, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        writer.writeheader()

        for resultado in resultados:
            writer.writerow(asdict(resultado))


def salvar_json(resultados: list[PortaResultado], arquivo_json: str) -> None:
    with open(arquivo_json, "w", encoding="utf-8") as f:
        json.dump([asdict(r) for r in resultados], f, ensure_ascii=False, indent=2)


# =========================================================
# CLI / MAIN
# =========================================================

def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Coleta status de portas fisicas de switches Cisco, Fortinet e Huawei."
    )

    parser.add_argument(
        "-i",
        "--inventario",
        default="switches.csv",
        help="Arquivo CSV de inventario (default: switches.csv)",
    )

    parser.add_argument(
        "-o",
        "--csv-saida",
        default="resultado_portas_fisicas.csv",
        help="Arquivo CSV de saida (default: resultado_portas_fisicas.csv)",
    )

    parser.add_argument(
        "-j",
        "--json-saida",
        default="dados_portas.json",
        help="Arquivo JSON de saida (default: dados_portas.json)",
    )

    parser.add_argument(
        "-w",
        "--workers",
        type=int,
        default=5,
        help="Numero maximo de conexoes simultaneas (default: 5)",
    )

    parser.add_argument(
        "-t",
        "--timeout",
        type=int,
        default=20,
        help="Timeout de conexao em segundos (default: 20)",
    )

    parser.add_argument(
        "-r",
        "--retries",
        type=int,
        default=3,
        help="Numero de tentativas por switch em caso de timeout (default: 3)",
    )

    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Logging detalhado (DEBUG).",
    )

    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    configurar_logging(args.verbose)

    try:
        switches = carregar_switches(args.inventario)
    except Exception as e:
        LOG.error("Falha ao carregar inventario: %s", e)
        return 1

    LOG.info("Switches carregados: %s", len(switches))
    LOG.info("Workers simultaneos: %s", args.workers)

    todos_resultados: list[PortaResultado] = []
    falhas = 0

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(
                coletar_switch,
                sw,
                args.timeout,
                args.retries,
            ): sw
            for sw in switches
        }

        for future in as_completed(futures):
            sw = futures[future]

            try:
                resultados = future.result()
                todos_resultados.extend(resultados)

            except Exception as e:
                falhas += 1
                LOG.error("Falha final em %s (%s): %s", sw.name, sw.host, e)

    try:
        salvar_csv(todos_resultados, args.csv_saida)
        salvar_json(todos_resultados, args.json_saida)
    except Exception as e:
        LOG.error("Falha ao salvar resultados: %s", e)
        return 1

    LOG.info("Coleta finalizada")
    LOG.info("Portas coletadas: %s", len(todos_resultados))
    LOG.info("Switches com falha: %s", falhas)
    LOG.info("CSV gerado: %s", args.csv_saida)
    LOG.info("JSON gerado: %s", args.json_saida)
    LOG.info("Log gerado: coleta_portas.log")

    return 0 if falhas == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
