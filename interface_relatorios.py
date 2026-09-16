"""Interface grafica para gerar relatorios de portas a partir do historico JSON.

A interface nao executa a coleta. O coletor continua sendo executado pelo
Agendador de Tarefas do Windows.

Arquivos esperados na mesma pasta:
- gerar_relatorio.py
- historico/
- relatorios/ (criada automaticamente)
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


PASTA_PROJETO = Path(__file__).resolve().parent
PASTA_HISTORICO = PASTA_PROJETO / "historico"
PASTA_RELATORIOS = PASTA_PROJETO / "relatorios"
PASTA_LOGS = PASTA_PROJETO / "logs"
SCRIPT_RELATORIO = PASTA_PROJETO / "gerar_relatorio.py"


class InterfaceRelatorios(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Sistema de Relatorios de Portas")
        self.geometry("820x590")
        self.minsize(760, 540)
        self.configure(bg="#eef2f7")

        self.tipo_relatorio = tk.StringVar(value="data")
        self.data_especifica = tk.StringVar()
        self.data_inicial = tk.StringVar()
        self.data_final = tk.StringVar()
        self.status = tk.StringVar(value="Selecione o tipo de relatorio.")
        self.ultimo_relatorio: Path | None = None

        PASTA_RELATORIOS.mkdir(parents=True, exist_ok=True)
        self._configurar_estilo()
        self._montar_tela()
        self.atualizar_datas()

    def _configurar_estilo(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass

        style.configure("Titulo.TLabel", font=("Segoe UI", 20, "bold"), foreground="#000000", background="#ffffff")
        style.configure("Subtitulo.TLabel", font=("Segoe UI", 10), foreground="#616972", background="#ffffff")
        style.configure("Secao.TLabelframe", background="#ffffff", borderwidth=1, relief="solid")
        style.configure("Secao.TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground="#000000", background="#ffffff")
        style.configure("TLabel", font=("Segoe UI", 10), background="#ffffff")
        style.configure("TRadiobutton", font=("Segoe UI", 10), background="#ffffff")
        style.configure("Acao.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        style.configure("TButton", font=("Segoe UI", 9), padding=(10, 6))

    def _montar_tela(self) -> None:
        cabecalho = ttk.Frame(self, padding=(24, 20, 24, 10))
        cabecalho.pack(fill="x")
        cabecalho.configure(style="TFrame")

        ttk.Label(cabecalho, text="Relatorios de Portas", style="Titulo.TLabel").pack(anchor="w")
        ttk.Label(
            cabecalho,
            text="Selecione uma data, um periodo ou todo o historico e gere o documento Word.",
            style="Subtitulo.TLabel",
        ).pack(anchor="w", pady=(3, 0))

        conteudo = ttk.Frame(self, padding=(24, 8, 24, 20))
        conteudo.pack(fill="both", expand=True)

        selecao = ttk.LabelFrame(conteudo, text=" Tipo de relatorio ", style="Secao.TLabelframe", padding=18)
        selecao.pack(fill="x")

        radios = ttk.Frame(selecao)
        radios.pack(fill="x")
        ttk.Radiobutton(radios, text="Data especifica", value="data", variable=self.tipo_relatorio, command=self._atualizar_campos).pack(side="left", padx=(0, 24))
        ttk.Radiobutton(radios, text="Periodo", value="periodo", variable=self.tipo_relatorio, command=self._atualizar_campos).pack(side="left", padx=(0, 24))
        ttk.Radiobutton(radios, text="Todo o historico", value="tudo", variable=self.tipo_relatorio, command=self._atualizar_campos).pack(side="left")

        campos = ttk.Frame(selecao)
        campos.pack(fill="x", pady=(18, 0))
        campos.columnconfigure(1, weight=1)
        campos.columnconfigure(3, weight=1)

        ttk.Label(campos, text="Data especifica:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.combo_data = ttk.Combobox(campos, textvariable=self.data_especifica, state="readonly", width=18)
        self.combo_data.grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=5)

        ttk.Label(campos, text="Data inicial:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        self.combo_inicio = ttk.Combobox(campos, textvariable=self.data_inicial, state="disabled", width=18)
        self.combo_inicio.grid(row=1, column=1, sticky="ew", padx=(0, 20), pady=5)

        ttk.Label(campos, text="Data final:").grid(row=1, column=2, sticky="w", padx=(0, 8), pady=5)
        self.combo_fim = ttk.Combobox(campos, textvariable=self.data_final, state="disabled", width=18)
        self.combo_fim.grid(row=1, column=3, sticky="ew", pady=5)

        ttk.Button(campos, text="Atualizar datas", command=self.atualizar_datas).grid(row=0, column=3, sticky="e", pady=5)

        acoes = ttk.LabelFrame(conteudo, text=" Relatorio ", style="Secao.TLabelframe", padding=18)
        acoes.pack(fill="x", pady=(14, 0))

        ttk.Button(acoes, text="Gerar relatorio", style="Acao.TButton", command=self.gerar_relatorio).pack(side="left", padx=(0, 10))
        self.btn_abrir = ttk.Button(acoes, text="Abrir ultimo relatorio", command=self.abrir_ultimo_relatorio, state="disabled")
        self.btn_abrir.pack(side="left", padx=(0, 10))
        ttk.Button(acoes, text="Abrir pasta de relatorios", command=lambda: self.abrir_pasta(PASTA_RELATORIOS)).pack(side="left")

        pastas = ttk.LabelFrame(conteudo, text=" Pastas do projeto ", style="Secao.TLabelframe", padding=18)
        pastas.pack(fill="x", pady=(14, 0))
        ttk.Button(pastas, text="Abrir historico", command=lambda: self.abrir_pasta(PASTA_HISTORICO)).pack(side="left", padx=(0, 10))
        ttk.Button(pastas, text="Abrir logs", command=lambda: self.abrir_pasta(PASTA_LOGS)).pack(side="left", padx=(0, 10))
        ttk.Button(pastas, text="Selecionar pasta historico", command=self.selecionar_historico).pack(side="left")

        status_frame = ttk.LabelFrame(conteudo, text=" Status ", style="Secao.TLabelframe", padding=14)
        status_frame.pack(fill="both", expand=True, pady=(14, 0))

        self.barra = ttk.Progressbar(status_frame, mode="indeterminate")
        self.barra.pack(fill="x", pady=(0, 10))
        self.label_status = ttk.Label(status_frame, textvariable=self.status, wraplength=720, justify="left")
        self.label_status.pack(fill="x", anchor="w")

    def _atualizar_campos(self) -> None:
        tipo = self.tipo_relatorio.get()
        self.combo_data.configure(state="readonly" if tipo == "data" else "disabled")
        estado_periodo = "readonly" if tipo == "periodo" else "disabled"
        self.combo_inicio.configure(state=estado_periodo)
        self.combo_fim.configure(state=estado_periodo)

    @staticmethod
    def _ler_datas_json(pasta: Path) -> list[str]:
        datas: set[str] = set()
        if not pasta.exists():
            return []

        for arquivo in pasta.rglob("portas_*.json"):
            try:
                with arquivo.open("r", encoding="utf-8") as f:
                    registros = json.load(f)
                if not isinstance(registros, list):
                    continue
                for item in registros:
                    valor = item.get("data_hora")
                    if valor:
                        data = datetime.fromisoformat(str(valor)).date()
                        datas.add(data.isoformat())
            except (OSError, ValueError, TypeError, json.JSONDecodeError):
                continue

        return sorted(datas)

    def atualizar_datas(self) -> None:
        datas = self._ler_datas_json(PASTA_HISTORICO)
        self.combo_data["values"] = datas
        self.combo_inicio["values"] = datas
        self.combo_fim["values"] = datas

        if datas:
            self.data_especifica.set(datas[-1])
            self.data_inicial.set(datas[0])
            self.data_final.set(datas[-1])
            self.status.set(f"{len(datas)} data(s) encontrada(s). Ultima data disponivel: {datas[-1]}.")
        else:
            self.data_especifica.set("")
            self.data_inicial.set("")
            self.data_final.set("")
            self.status.set(f"Nenhuma coleta encontrada em: {PASTA_HISTORICO}")

    def selecionar_historico(self) -> None:
        global PASTA_HISTORICO
        selecionada = filedialog.askdirectory(title="Selecione a pasta do historico", initialdir=PASTA_HISTORICO)
        if selecionada:
            PASTA_HISTORICO = Path(selecionada)
            self.atualizar_datas()

    def _montar_comando(self) -> tuple[list[str], Path]:
        if not SCRIPT_RELATORIO.exists():
            raise FileNotFoundError(f"Gerador nao encontrado: {SCRIPT_RELATORIO}")
        if not PASTA_HISTORICO.exists():
            raise FileNotFoundError(f"Pasta de historico nao encontrada: {PASTA_HISTORICO}")

        tipo = self.tipo_relatorio.get()
        comando = [sys.executable, str(SCRIPT_RELATORIO), "-i", str(PASTA_HISTORICO)]

        if tipo == "data":
            data = self.data_especifica.get().strip()
            if not data:
                raise ValueError("Selecione uma data especifica.")
            nome = f"Relatorio_Portas_{data}.docx"
            comando += ["--data", data]

        elif tipo == "periodo":
            inicio = self.data_inicial.get().strip()
            fim = self.data_final.get().strip()
            if not inicio or not fim:
                raise ValueError("Selecione a data inicial e a data final.")
            if inicio > fim:
                raise ValueError("A data inicial nao pode ser posterior a data final.")
            nome = f"Relatorio_Portas_{inicio}_a_{fim}.docx"
            comando += ["--inicio", inicio, "--fim", fim]

        else:
            nome = f"Relatorio_Portas_Completo_{datetime.now():%Y%m%d_%H%M%S}.docx"

        saida = PASTA_RELATORIOS / nome
        comando += ["-o", str(saida)]
        return comando, saida

    def gerar_relatorio(self) -> None:
        try:
            comando, saida = self._montar_comando()
        except Exception as erro:
            messagebox.showerror("Nao foi possivel gerar", str(erro))
            return

        self.status.set("Gerando relatorio. Aguarde...")
        self.barra.start(10)
        self._bloquear_interface(True)

        thread = threading.Thread(target=self._executar_gerador, args=(comando, saida), daemon=True)
        thread.start()

    def _executar_gerador(self, comando: list[str], saida: Path) -> None:
        try:
            processo = subprocess.run(
                comando,
                cwd=PASTA_PROJETO,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0,
            )
            if processo.returncode != 0:
                detalhe = (processo.stdout + "\n" + processo.stderr).strip()
                raise RuntimeError(detalhe or f"O gerador terminou com codigo {processo.returncode}.")
            if not saida.exists():
                raise RuntimeError("O processo terminou, mas o arquivo DOCX nao foi criado.")
            self.after(0, self._geracao_concluida, saida, processo.stdout.strip())
        except Exception as erro:
            self.after(0, self._geracao_falhou, str(erro))

    def _geracao_concluida(self, saida: Path, detalhe: str) -> None:
        self.barra.stop()
        self._bloquear_interface(False)
        self.ultimo_relatorio = saida
        self.btn_abrir.configure(state="normal")
        self.status.set(f"Relatorio gerado com sucesso:\n{saida}\n\n{detalhe}")
        messagebox.showinfo("Relatorio gerado", f"Arquivo criado com sucesso:\n\n{saida}")

    def _geracao_falhou(self, erro: str) -> None:
        self.barra.stop()
        self._bloquear_interface(False)
        self.status.set(f"Erro ao gerar relatorio:\n{erro}")
        messagebox.showerror("Erro ao gerar relatorio", erro)

    def _bloquear_interface(self, bloquear: bool) -> None:
        cursor = "wait" if bloquear else ""
        self.configure(cursor=cursor)
        self.update_idletasks()

    def abrir_ultimo_relatorio(self) -> None:
        if not self.ultimo_relatorio or not self.ultimo_relatorio.exists():
            messagebox.showwarning("Relatorio", "Nenhum relatorio foi gerado nesta sessao.")
            return
        self._abrir_caminho(self.ultimo_relatorio)

    def abrir_pasta(self, pasta: Path) -> None:
        pasta.mkdir(parents=True, exist_ok=True)
        self._abrir_caminho(pasta)

    @staticmethod
    def _abrir_caminho(caminho: Path) -> None:
        try:
            if os.name == "nt":
                os.startfile(str(caminho))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(caminho)])
            else:
                subprocess.Popen(["xdg-open", str(caminho)])
        except Exception as erro:
            messagebox.showerror("Nao foi possivel abrir", str(erro))


if __name__ == "__main__":
    app = InterfaceRelatorios()
    app.mainloop()
