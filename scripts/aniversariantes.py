#!/usr/bin/env python3
"""Envia para um grupo do WhatsApp a lista de aniversariantes do dia.

Lê uma planilha (Excel .xlsx ou CSV) com as colunas abaixo — a ordem não
importa e acentos/maiúsculas são opcionais:

    Nome | Nascimento | Cidade | Cargo | Telefone

`Nascimento` aceita "25/12", "25/12/1980", "1980-12-25" ou uma data do Excel.

De onde vem a planilha (variável ANIVERSARIANTES_FONTE):
  * um link (http/https) — ex.: link de exportação de uma Google Sheet;
  * um caminho local     — ex.: data/aniversariantes.xlsx;
  * ou o próprio conteúdo CSV na variável ANIVERSARIANTES_CSV.

Como envia (variável WHATSAPP_PROVEDOR):
  * zapi      — Z-API   (ZAPI_INSTANCIA, ZAPI_TOKEN, ZAPI_CLIENT_TOKEN)
  * evolution — Evolution API (EVOLUTION_URL, EVOLUTION_INSTANCIA, EVOLUTION_APIKEY)
O grupo de destino vai em WHATSAPP_GRUPO_ID. Para testar mandando só para
você, defina WHATSAPP_TESTE_NUMERO (ex.: 5519999990000) e ENVIAR_SO_PARA_MIM=1.

Uso:
    python3 scripts/aniversariantes.py --dry-run            # só mostra a mensagem
    python3 scripts/aniversariantes.py --data 25/12 --dry-run
    python3 scripts/aniversariantes.py                       # envia de verdade
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
import unicodedata
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

FUSO = ZoneInfo(os.environ.get("FUSO_HORARIO", "America/Sao_Paulo"))

# No GitHub Actions os logs de repositórios públicos são públicos: nunca
# imprimir nomes ou telefones lá.
SILENCIOSO = os.environ.get("GITHUB_ACTIONS") == "true" \
    or os.environ.get("SILENCIOSO", "").lower() in ("1", "true", "sim")

# Nome normalizado da coluna -> campo interno
COLUNAS = {
    "nome": "nome",
    "nascimento": "nascimento",
    "data de nascimento": "nascimento",
    "aniversario": "nascimento",
    "data": "nascimento",
    "cidade": "cidade",
    "municipio": "cidade",
    "cargo": "cargo",
    "funcao": "cargo",
    "telefone": "telefone",
    "celular": "telefone",
    "whatsapp": "telefone",
}


@dataclass
class Pessoa:
    nome: str
    dia: int
    mes: int
    cidade: str = ""
    cargo: str = ""
    telefone: str = ""
    whatsapp: str = ""  # número no formato internacional, só dígitos (ex.: 5519999990000)


# --------------------------------------------------------------------------- #
# Leitura da planilha
# --------------------------------------------------------------------------- #

def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", texto).strip().lower()


def ler_bytes_fonte() -> tuple[bytes, str]:
    """Devolve (conteúdo, tipo) onde tipo é 'xlsx' ou 'csv'."""
    inline = os.environ.get("ANIVERSARIANTES_CSV", "").strip()
    if inline:
        return inline.encode("utf-8"), "csv"

    fonte = os.environ.get("ANIVERSARIANTES_FONTE", "").strip()
    if not fonte:
        sys.exit("Defina ANIVERSARIANTES_FONTE (link ou caminho) ou ANIVERSARIANTES_CSV.")

    if fonte.startswith(("http://", "https://")):
        req = urllib.request.Request(fonte, headers={"User-Agent": "aniversariantes-bot"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            conteudo = resp.read()
            tipo_http = resp.headers.get("Content-Type", "")
        if conteudo[:2] == b"PK" or "spreadsheetml" in tipo_http:
            return conteudo, "xlsx"
        return conteudo, "csv"

    caminho = Path(fonte)
    if not caminho.exists():
        sys.exit(f"Planilha não encontrada: {caminho}")
    tipo = "xlsx" if caminho.suffix.lower() in (".xlsx", ".xlsm") else "csv"
    return caminho.read_bytes(), tipo


def linhas_xlsx(conteudo: bytes) -> list[list]:
    from openpyxl import load_workbook

    wb = load_workbook(io.BytesIO(conteudo), read_only=True, data_only=True)
    ws = wb.active
    return [list(linha) for linha in ws.iter_rows(values_only=True)]


def linhas_csv(conteudo: bytes) -> list[list]:
    texto = conteudo.decode("utf-8-sig", errors="replace")
    try:
        dialeto = csv.Sniffer().sniff(texto[:2048], delimiters=",;\t")
    except csv.Error:
        dialeto = csv.excel
    return [linha for linha in csv.reader(io.StringIO(texto), dialeto)]


def extrair_dia_mes(valor) -> tuple[int, int] | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, (datetime, date)):
        return valor.day, valor.month
    texto = str(valor).strip()
    # 1980-12-25 (ISO)
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", texto)
    if m:
        return int(m.group(3)), int(m.group(2))
    # 25/12, 25/12/1980, 25-12, 25.12.1980
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})(?:[/\-.]\d{2,4})?$", texto)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def carregar_pessoas() -> list[Pessoa]:
    conteudo, tipo = ler_bytes_fonte()
    linhas = linhas_xlsx(conteudo) if tipo == "xlsx" else linhas_csv(conteudo)
    linhas = [l for l in linhas if any(c not in (None, "") for c in l)]
    if not linhas:
        sys.exit("A planilha está vazia.")

    cabecalho = [COLUNAS.get(normalizar(c or "")) for c in linhas[0]]
    for obrigatoria in ("nome", "nascimento"):
        if obrigatoria not in cabecalho:
            sys.exit(f"Coluna obrigatória ausente na planilha: {obrigatoria!r}. "
                     f"Cabeçalho lido: {linhas[0]}")

    pessoas: list[Pessoa] = []
    for numero, linha in enumerate(linhas[1:], start=2):
        dados = {}
        for campo, valor in zip(cabecalho, linha):
            if campo and campo not in dados:
                dados[campo] = valor
        nome = str(dados.get("nome") or "").strip()
        if not nome:
            continue
        dia_mes = extrair_dia_mes(dados.get("nascimento"))
        if not dia_mes:
            quem = "" if SILENCIOSO else f" ({nome})"
            print(f"Aviso: linha {numero}{quem} sem data de nascimento válida — ignorada.",
                  file=sys.stderr)
            continue
        pessoas.append(Pessoa(
            nome=nome,
            dia=dia_mes[0],
            mes=dia_mes[1],
            cidade=str(dados.get("cidade") or "").strip(),
            cargo=str(dados.get("cargo") or "").strip(),
            telefone=formatar_telefone(dados.get("telefone")),
            whatsapp=numero_whatsapp(dados.get("telefone")),
        ))
    return pessoas


# --------------------------------------------------------------------------- #
# Regras e mensagem
# --------------------------------------------------------------------------- #

def eh_bissexto(ano: int) -> bool:
    return ano % 4 == 0 and (ano % 100 != 0 or ano % 400 == 0)


def aniversariantes_do_dia(pessoas: list[Pessoa], hoje: date) -> list[Pessoa]:
    def faz_hoje(p: Pessoa) -> bool:
        if (p.dia, p.mes) == (hoje.day, hoje.month):
            return True
        # Nascidos em 29/02 comemoram em 28/02 nos anos não bissextos.
        return (p.dia, p.mes) == (29, 2) and (hoje.day, hoje.month) == (28, 2) \
            and not eh_bissexto(hoje.year)

    return sorted((p for p in pessoas if faz_hoje(p)), key=lambda p: normalizar(p.nome))


def so_digitos(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\D", "", str(valor))


def numero_whatsapp(valor) -> str:
    """Número para o link wa.me: DDI 55 + DDD + número, só dígitos."""
    digitos = so_digitos(valor).lstrip("0")
    if str(valor).strip().startswith("+"):   # já veio com DDI explícito
        return digitos if len(digitos) >= 8 else ""
    if len(digitos) in (10, 11):          # brasileiro sem DDI
        return "55" + digitos
    if digitos.startswith("55") and len(digitos) in (12, 13):
        return digitos
    return digitos if len(digitos) >= 10 else ""   # outro país, já com DDI


def formatar_telefone(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    digitos = so_digitos(valor).lstrip("0")
    if texto.startswith("+") and not texto.startswith("+55"):
        return texto                      # número estrangeiro: mantém como veio
    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos = digitos[2:]
    if len(digitos) == 11:
        return f"({digitos[:2]}) {digitos[2:7]}-{digitos[7:]}"
    if len(digitos) == 10:
        return f"({digitos[:2]}) {digitos[2:6]}-{digitos[6:]}"
    return str(valor).strip()


def montar_mensagem(lista: list[Pessoa], hoje: date) -> str:
    titulo = os.environ.get("MENSAGEM_TITULO") or "🎉 *Aniversariantes de hoje*"
    partes = [f"{titulo} ({hoje:%d/%m})", ""]
    for p in lista:
        bloco = [f"🎂 *{p.nome}*"]
        local = " — ".join(x for x in (p.cidade, p.cargo) if x)
        if local:
            bloco.append(f"📍 {local}")
        if p.telefone:
            bloco.append(f"📞 {p.telefone}")
        if p.whatsapp:
            # Link clicável que abre a conversa privada com o aniversariante.
            bloco.append(f"💬 Mensagem privada: https://wa.me/{p.whatsapp}")
        partes.append("\n".join(bloco))
        partes.append("")
    rodape = os.environ.get("MENSAGEM_RODAPE", "").strip()
    if rodape:
        partes.append(rodape)
    return "\n".join(partes).strip()


# --------------------------------------------------------------------------- #
# Envio
# --------------------------------------------------------------------------- #

def exigir(nome: str) -> str:
    valor = os.environ.get(nome, "").strip()
    if not valor:
        sys.exit(f"Variável obrigatória não definida: {nome}")
    return valor


def post_json(url: str, corpo: dict, cabecalhos: dict) -> str:
    req = urllib.request.Request(
        url,
        data=json.dumps(corpo).encode("utf-8"),
        headers={"Content-Type": "application/json", **cabecalhos},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as erro:
        detalhe = erro.read().decode("utf-8", errors="replace")
        sys.exit(f"Falha no envio (HTTP {erro.code}): {detalhe[:500]}")


def enviar(mensagem: str) -> None:
    provedor = (os.environ.get("WHATSAPP_PROVEDOR") or "zapi").strip().lower()
    if os.environ.get("ENVIAR_SO_PARA_MIM", "").lower() in ("1", "true", "sim"):
        grupo = numero_whatsapp(exigir("WHATSAPP_TESTE_NUMERO"))
        if not grupo:
            sys.exit("WHATSAPP_TESTE_NUMERO inválido (use DDI + DDD + número).")
        mensagem = "🧪 *TESTE — enviado só para você*\n\n" + mensagem
        print("Modo 'só para mim': enviando para o número de teste, NÃO para o grupo.")
    else:
        grupo = exigir("WHATSAPP_GRUPO_ID")

    if provedor == "zapi":
        instancia = exigir("ZAPI_INSTANCIA")
        token = exigir("ZAPI_TOKEN")
        cabecalhos = {}
        if os.environ.get("ZAPI_CLIENT_TOKEN"):
            cabecalhos["Client-Token"] = os.environ["ZAPI_CLIENT_TOKEN"]
        url = f"https://api.z-api.io/instances/{instancia}/token/{token}/send-text"
        resposta = post_json(url, {"phone": grupo, "message": mensagem}, cabecalhos)

    elif provedor == "evolution":
        base = exigir("EVOLUTION_URL").rstrip("/")
        instancia = exigir("EVOLUTION_INSTANCIA")
        apikey = exigir("EVOLUTION_APIKEY")
        # IDs de grupo têm mais de 15 dígitos (ou hífen); telefones não.
        eh_grupo = "-" in grupo or len(re.sub(r"\D", "", grupo)) > 15
        numero = grupo if "@" in grupo or not eh_grupo else f"{grupo}@g.us"
        url = f"{base}/message/sendText/{instancia}"
        resposta = post_json(url, {"number": numero, "text": mensagem, "linkPreview": False}, {"apikey": apikey})

    else:
        sys.exit(f"WHATSAPP_PROVEDOR desconhecido: {provedor!r} (use 'zapi' ou 'evolution').")

    print(f"Mensagem enviada via {provedor}.")
    if not SILENCIOSO:
        print(f"Resposta: {resposta[:300]}")


# --------------------------------------------------------------------------- #

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--dry-run", action="store_true",
                        help="mostra a mensagem sem enviar")
    parser.add_argument("--data", help="simula outro dia (dd/mm ou dd/mm/aaaa)")
    args = parser.parse_args()

    dry_run = args.dry_run or os.environ.get("DRY_RUN", "").lower() in ("1", "true", "sim")

    hoje = datetime.now(FUSO).date()
    if args.data:
        partes = [int(x) for x in re.split(r"[/\-.]", args.data)]
        hoje = date(partes[2] if len(partes) > 2 else hoje.year, partes[1], partes[0])

    pessoas = carregar_pessoas()
    lista = aniversariantes_do_dia(pessoas, hoje)
    print(f"{len(pessoas)} pessoas na planilha; {len(lista)} aniversariante(s) em {hoje:%d/%m/%Y}.")

    if not lista:
        aviso = os.environ.get("MENSAGEM_SEM_ANIVERSARIANTES", "").strip()
        if not aviso:
            print("Nenhum aniversariante hoje — nada a enviar.")
            return
        mensagem = aviso
    else:
        mensagem = montar_mensagem(lista, hoje)

    if SILENCIOSO:
        print("(conteúdo da mensagem oculto no log para proteger dados pessoais)")
    else:
        print("-" * 40)
        print(mensagem)
        print("-" * 40)

    if dry_run:
        print("Modo teste (dry-run): mensagem NÃO enviada.")
        return
    enviar(mensagem)


if __name__ == "__main__":
    main()
