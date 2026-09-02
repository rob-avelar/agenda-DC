#!/usr/bin/env python3
"""Gera a planilha de exemplo em data/agenda.xlsx.

Colunas esperadas pelo app (a ordem não importa, os nomes sim — acentos são
opcionais):

    Dia | Data | Início | Fim | Atividade | Local | Responsável | Categoria | Observações

Categorias com cor própria no app: Reunião, Treinamento, Evento, Pausa, Tarefa.
Qualquer outra categoria recebe a cor neutra.

Uso:
    python3 scripts/gerar_planilha_exemplo.py
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

CABECALHOS = [
    "Dia", "Data", "Início", "Fim", "Atividade",
    "Local", "Responsável", "Categoria", "Observações",
]

# (Dia, Data, Início, Fim, Atividade, Local, Responsável, Categoria, Observações)
LINHAS = [
    ("Segunda-feira", "08/09", "09:00", "09:30", "Abertura da semana", "Auditório", "Robson", "Reunião", "Alinhamento geral"),
    ("Segunda-feira", "08/09", "09:30", "11:00", "Planejamento de sprint", "Sala 1", "Equipe", "Reunião", ""),
    ("Segunda-feira", "08/09", "11:00", "12:00", "Treinamento de segurança", "Sala 2", "Ana", "Treinamento", "Obrigatório"),
    ("Segunda-feira", "08/09", "12:00", "13:00", "Almoço", "Refeitório", "", "Pausa", ""),
    ("Segunda-feira", "08/09", "14:00", "17:00", "Execução de tarefas", "Open space", "Equipe", "Tarefa", ""),

    ("Terça-feira", "09/09", "09:00", "10:30", "Revisão de indicadores", "Sala 1", "Robson", "Reunião", "KPIs do mês"),
    ("Terça-feira", "09/09", "10:30", "12:00", "Workshop de processos", "Sala 3", "Carlos", "Treinamento", ""),
    ("Terça-feira", "09/09", "12:00", "13:00", "Almoço", "Refeitório", "", "Pausa", ""),
    ("Terça-feira", "09/09", "13:00", "15:00", "Atendimento a clientes", "Recepção", "Bruna", "Tarefa", ""),

    ("Quarta-feira", "10/09", "08:30", "09:00", "Café da manhã coletivo", "Copa", "", "Pausa", ""),
    ("Quarta-feira", "10/09", "09:00", "12:00", "Auditoria interna", "Sala 1", "Auditoria", "Evento", "Duração estimada"),
    ("Quarta-feira", "10/09", "14:00", "16:00", "Reunião com fornecedores", "Sala 2", "Robson", "Reunião", ""),

    ("Quinta-feira", "11/09", "09:00", "11:00", "Treinamento de sistema", "Lab", "Ana", "Treinamento", "Trazer notebook"),
    ("Quinta-feira", "11/09", "11:00", "12:00", "Retrospectiva", "Sala 1", "Equipe", "Reunião", ""),
    ("Quinta-feira", "11/09", "12:00", "13:00", "Almoço", "Refeitório", "", "Pausa", ""),
    ("Quinta-feira", "11/09", "13:30", "17:00", "Desenvolvimento", "Open space", "Equipe", "Tarefa", ""),

    ("Sexta-feira", "12/09", "09:00", "10:00", "Planejamento da próxima semana", "Sala 1", "Robson", "Reunião", ""),
    ("Sexta-feira", "12/09", "10:00", "12:00", "Confraternização", "Área externa", "Todos", "Evento", "Encerramento da semana"),
    ("Sexta-feira", "12/09", "12:00", "13:00", "Almoço", "Refeitório", "", "Pausa", ""),
]


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Programação"

    header_fill = PatternFill("solid", fgColor="1E293B")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append(CABECALHOS)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for linha in LINHAS:
        ws.append(list(linha))

    larguras = [16, 8, 9, 9, 32, 16, 16, 14, 24]
    for i, largura in enumerate(larguras, start=1):
        ws.column_dimensions[chr(64 + i)].width = largura

    ws.freeze_panes = "A2"

    destino = Path(__file__).resolve().parent.parent / "data" / "agenda.xlsx"
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    print(f"Planilha gerada: {destino} ({len(LINHAS)} atividades)")


if __name__ == "__main__":
    main()
