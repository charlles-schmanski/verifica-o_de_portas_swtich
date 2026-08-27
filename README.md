# Monitor de Portas Físicas

Projeto em Python para coletar o status das portas físicas de switches Cisco, Fortinet/FortiSwitch e Huawei usando Netmiko.

O script acessa os switches via SSH ou Telnet, executa comandos de consulta e gera os arquivos de resultado em CSV, JSON e log.

## Equipamentos suportados

cisco\_ios
cisco\_ios\_telnet
fortinet
huawei
huawei\_telnet



## Comandos usados

### Cisco



show interfaces status



### Fortinet/FortiSwitch



get switch physical-port all



### Huawei



display interface brief



## Instalação

Instale o Netmiko:



pip install netmiko



## Inventário

Crie o arquivo `switches.csv` na mesma pasta do script.

Modelo:

csv
name,host,device\_type,username,password,secret
SW-CISCO-01,10.10.10.1,cisco\_ios,admin,senha,enable
SW-CISCO-TELNET-01,10.10.10.2,cisco\_ios\_telnet,admin,senha,enable
FSW-01,10.10.10.3,fortinet,admin,senha,
HUA-SW-01,10.10.10.4,huawei,admin,senha,
HUA-SW-02,10.10.10.5,huawei\_telnet,admin,senha,



### Campos do inventário

* `name`: nome amigável do switch.
* `host`: IP ou DNS do switch.
* `device\\\_type`: tipo do equipamento no Netmiko.
* `username`: usuário de acesso.
* `password`: senha de acesso.
* `secret`: senha de enable, normalmente usada em Cisco. Para Fortinet e Huawei pode ficar vazio.

## Como executar

Execução simples:

```bash
python coleta\\\_portas.py
```

Execução recomendada:

```bash
python coleta\\\_portas.py -i switches.csv -w 5 -t 10 -r 2
```

Exemplo completo:

```bash
python coleta\\\_portas.py -i switches.csv -o resultado\\\_portas\\\_fisicas.csv -j dados\\\_portas.json -w 5 -t 10 -r 2
```

## Parâmetros principais

```text
-i  arquivo de inventário
-o  arquivo CSV de saída
-j  arquivo JSON de saída
-w  conexões simultâneas
-t  timeout de conexão
-r  tentativas em caso de timeout
-v  modo detalhado, com mais logs
```

## Arquivos gerados

Após a execução, o script gera:

```text
resultado\\\_portas\\\_fisicas.csv
dados\\\_portas.json
coleta\\\_portas.log
```

### `resultado\\\_portas\\\_fisicas.csv`

Arquivo em formato CSV para abrir no Excel ou usar em relatórios.

### `dados\\\_portas.json`

Arquivo usado pelo dashboard HTML separado.

### `coleta\\\_portas.log`

Arquivo de log da execução, com sucesso, erro, timeout e autenticação.

## Dashboard HTML

Depois de rodar a coleta, será gerado:

```text
dados\\\_portas.json
```

Para abrir o dashboard local:

```bash
python -m http.server 8000
```

Depois acesse no navegador:

```text
http://localhost:8000/dashboard.html
```

## Status normalizados

```text
UP          porta ativa
DOWN        porta sem link ou inativa
ADMIN\\\_DOWN  porta desligada administrativamente
UNKNOWN     status não identificado
```

## Segurança

O script executa apenas comandos de consulta.

Ele não altera configuração dos switches.

Recomendações:

* Proteja o arquivo `switches.csv`, pois ele contém credenciais.
* Use uma conta com privilégio mínimo necessário.
* Use `-w 5` para limitar conexões simultâneas.
* Teste primeiro com poucos switches antes de executar em todos.

## Teste inicial recomendado

Antes de rodar em todos os equipamentos, teste com apenas alguns switches no `switches.csv` e execute:

```bash
python coleta\\\_portas\\\_corrigido.py -i switches.csv -w 1 -v
```

Se tudo funcionar, rode com mais conexões simultâneas:

```bash
python coleta\\\_portas\\\_corrigido.py -i switches.csv -w 5
```

## Estrutura simples do projeto

```text
monitor\\\_portas/
├── coleta\\\_portas\\\_corrigido.py
├── dashboard.html
├── switches.csv
├── requirements.txt
└── README.md
```

