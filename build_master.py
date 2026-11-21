#!/usr/bin/env python3
"""Gera a planilha no formato da Timesheet MASTER a partir do painel gerado por build.py.

Uso:
    python3 build_master.py <Painel.html> <saida.xlsx>

Layout (igual à MASTER): uma aba por mês; para cada cliente um bloco
"TIMESHEET - CLIENTE" com Code | Name | Job Type | Date | Time Spent | Hourly Charge | Total | Staff Name,
uma linha por cliente × serviço × pessoa no mês, e o total do cliente com a situação (Invoiced / To invoice).
Hourly Charge vem da aba Rates (preço por tipo de serviço); Total = Time Spent × Hourly Charge.
Só entram horas de clientes; tempo interno, almoço e férias ficam de fora.
"""
import json
import re
import sys
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July", "August",
             "September", "October", "November", "December"]
NOCLIENT = "Vários clientes (sem nome)"
F = "Arial"
BOLD = Font(name=F, bold=True)
TITLE = Font(name=F, bold=True, underline="single", size=12)
HDR = Font(name=F, bold=True, color="FFFFFF")
BODY = Font(name=F)
NOTE = Font(name=F, italic=True, color="666666", size=9)
INPUT = PatternFill("solid", fgColor="FFFF00")
HDR_FILL = PatternFill("solid", fgColor="0F6E5C")
TOT_FILL = PatternFill("solid", fgColor="D8ECE7")
THIN = Side(style="thin", color="BBBBBB")
GBP = '£#,##0.00;-£#,##0.00;"-"'
HRS = '0.00'
# Preços que já constam na Timesheet MASTER (outubro, Noris Medical): VAT £35/h e CT £100/h.
KNOWN_RATES = {"VAT": 35, "Accounts & CT": 100}


def load_data(html_path):
    s = open(html_path, encoding="utf-8").read()
    m = re.search(r"const DATA = (\{.*?\});\n</script>", s, re.S)
    if not m:
        sys.exit("DATA não encontrado no painel")
    return json.loads(m.group(1).replace("<\\/", "</"))


def main(html_path, out_path):
    D = load_data(html_path)
    rows = [dict(staff=D["staff"][r[0]], date=r[1], client=D["clients"][r[2]], service=D["services"][r[3]], hours=r[4])
            for r in D["rows"] if r[5] == "c"]

    agg = defaultdict(lambda: {"hours": 0.0, "last": ""})
    for r in rows:
        k = (r["date"][:7], r["client"], r["service"], r["staff"])
        agg[k]["hours"] += r["hours"]
        agg[k]["last"] = max(agg[k]["last"], r["date"])

    months = sorted({k[0] for k in agg})
    services = sorted({k[2] for k in agg})
    staff = sorted({k[3] for k in agg})
    wb = Workbook()

    # ---------- Rates ----------
    ws = wb.active
    ws.title = "Rates"
    ws["A1"] = "HOURLY CHARGE PER JOB TYPE / PREÇO POR HORA DE CADA SERVIÇO"
    ws["A1"].font = TITLE
    ws["A2"] = ("Fill in the yellow cells; every month tab picks the rate up automatically. "
                "Preencha as células amarelas; todas as abas de mês usam esse preço automaticamente.")
    ws["A2"].font = NOTE
    ws.append([])
    ws.append(["Job Type", "Hourly Charge (£)", "Client hours this year"])
    for c in ws[4]:
        c.font, c.fill = HDR, HDR_FILL
    hours_by_service = defaultdict(float)
    for k, v in agg.items():
        hours_by_service[k[2]] += v["hours"]
    for i, sname in enumerate(sorted(services, key=lambda x: -hours_by_service[x]), start=5):
        ws.cell(i, 1, sname).font = BODY
        c = ws.cell(i, 2, KNOWN_RATES.get(sname))
        c.fill, c.font, c.number_format = INPUT, Font(name=F, color="0000FF"), GBP
        if sname in KNOWN_RATES:
            c.comment = Comment("Valor já usado na Timesheet MASTER (outubro, Noris Medical). Confirme.", "Timesheet")
        ws.cell(i, 3, round(hours_by_service[sname], 2)).number_format = HRS
        ws.cell(i, 3).font = BODY
    last_rate_row = 4 + len(services)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 24
    rate_range = f"Rates!$A$5:$B${last_rate_row}"

    # ---------- Month tabs (MASTER layout) ----------
    status_dv_by_ws = {}
    for mk in months:
        y, m = map(int, mk.split("-"))
        ws = wb.create_sheet(MONTHS_EN[m - 1] if len({x[:4] for x in months}) == 1 else f"{MONTHS_EN[m-1]} {y}")
        dv = DataValidation(type="list", formula1='"Invoiced,To invoice,Not billable"', allow_blank=True)
        ws.add_data_validation(dv)
        status_dv_by_ws[ws.title] = dv
        widths = {"A": 3, "B": 14, "C": 34, "D": 24, "E": 12, "F": 14, "G": 15, "H": 14, "I": 16}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        ws["B1"] = f"{MONTHS_EN[m-1].upper()} {y} — client hours from the staff timesheets (decimal hours: 1.50 = 1h30)"
        ws["B1"].font = NOTE

        by_client = defaultdict(list)
        for k, v in agg.items():
            if k[0] == mk:
                by_client[k[1]].append((k[2], k[3], v))
        order = sorted((c for c in by_client if c != NOCLIENT), key=lambda c: -sum(x[2]["hours"] for x in by_client[c]))
        if NOCLIENT in by_client:
            order.append(NOCLIENT)

        r = 3
        for client in order:
            ws.cell(r, 2, f"TIMESHEET - {client.upper()}").font = TITLE
            r += 1
            for j, h in enumerate(["Code", "Name", "Job Type", "Date", "Time Spent", "Hourly Charge", "Total", "Staff Name"]):
                c = ws.cell(r, 2 + j, h)
                c.font, c.fill = HDR, HDR_FILL
                c.alignment = Alignment(horizontal="center")
            r += 1
            first = r
            for service, person, v in sorted(by_client[client], key=lambda x: (x[0], x[1])):
                ws.cell(r, 3, client).font = BODY
                ws.cell(r, 4, service).font = BODY
                c = ws.cell(r, 5, date.fromisoformat(v["last"]))
                c.number_format, c.font = "dd/mm/yyyy", BODY
                c = ws.cell(r, 6, round(v["hours"], 2))
                c.number_format, c.font = HRS, BODY
                c = ws.cell(r, 7, f"=IFERROR(INDEX(Rates!$B$5:$B${last_rate_row},MATCH(D{r},Rates!$A$5:$A${last_rate_row},0)),0)")
                c.number_format, c.font = GBP, Font(name=F, color="008000")
                c = ws.cell(r, 8, f"=F{r}*G{r}")
                c.number_format, c.font = GBP, BODY
                ws.cell(r, 9, person).font = BODY
                for col in range(2, 10):
                    ws.cell(r, col).border = Border(bottom=THIN)
                r += 1
            # total do cliente + situação (como "£ 1,502.25 | Invoiced" na MASTER)
            ws.cell(r, 5, "Total").font = BOLD
            c = ws.cell(r, 6, f"=SUM(F{first}:F{r-1})")
            c.number_format, c.font = HRS, BOLD
            c = ws.cell(r, 8, f"=SUM(H{first}:H{r-1})")
            c.number_format, c.font = GBP, BOLD
            st = ws.cell(r, 9)
            st.fill, st.font = INPUT, BODY
            status_dv_by_ws[ws.title].add(st)
            for col in range(5, 10):
                if col != 9:
                    ws.cell(r, col).fill = TOT_FILL
            r += 3
        ws.freeze_panes = "B3"

    # ---------- Summary ----------
    ws = wb.create_sheet("Summary", 1)
    ws["A1"] = "SUMMARY / RESUMO"
    ws["A1"].font = TITLE
    ws["A2"] = ("Client hours and £ per month and per person, calculated from the month tabs. Internal time, lunch and holidays are not included. "
                "Horas de clientes e £ por mês e por pessoa, calculados das abas de mês. Tempo interno, almoço e férias não entram.")
    ws["A2"].font = NOTE
    month_tabs = [s.title for s in wb.worksheets if s.title not in ("Rates", "Summary")]
    ws.append([])
    ws.append(["Month", "Client hours", "Total (£)", "£ per client hour"])
    for c in ws[4]:
        c.font, c.fill = HDR, HDR_FILL
    for i, t in enumerate(month_tabs, start=5):
        q = f"'{t}'" if " " in t else t
        ws.cell(i, 1, t).font = BODY
        c = ws.cell(i, 2, f'=SUMIF({q}!$I:$I,"<>",{q}!$F:$F)')
        c.number_format, c.font = HRS, Font(name=F, color="008000")
        c = ws.cell(i, 3, f'=SUMIF({q}!$I:$I,"<>",{q}!$H:$H)')
        c.number_format, c.font = GBP, Font(name=F, color="008000")
        c = ws.cell(i, 4, f"=IF(B{i}=0,0,C{i}/B{i})")
        c.number_format, c.font = GBP, BODY
    tr = 5 + len(month_tabs)
    ws.cell(tr, 1, "Total").font = BOLD
    for col, fmt in ((2, HRS), (3, GBP)):
        L = "ABCD"[col - 1]
        c = ws.cell(tr, col, f"=SUM({L}5:{L}{tr-1})")
        c.number_format, c.font = fmt, BOLD
    c = ws.cell(tr, 4, f"=IF(B{tr}=0,0,C{tr}/B{tr})")
    c.number_format, c.font = GBP, BOLD

    r0 = tr + 3
    ws.cell(r0 - 1, 1, "CLIENT HOURS PER PERSON / HORAS EM CLIENTES POR PESSOA").font = BOLD
    ws.cell(r0, 1, "Staff Name")
    for j, t in enumerate(month_tabs, start=2):
        ws.cell(r0, j, t)
    ws.cell(r0, len(month_tabs) + 2, "Total")
    for c in ws[r0]:
        c.font, c.fill = HDR, HDR_FILL
    for i, p in enumerate(staff, start=r0 + 1):
        ws.cell(i, 1, p).font = BODY
        for j, t in enumerate(month_tabs, start=2):
            q = f"'{t}'" if " " in t else t
            c = ws.cell(i, j, f'=SUMIF({q}!$I:$I,$A{i},{q}!$F:$F)')
            c.number_format, c.font = '0.00;-0.00;"-"', Font(name=F, color="008000")
        L = ws.cell(i, len(month_tabs) + 1).column_letter
        c = ws.cell(i, len(month_tabs) + 2, f"=SUM(B{i}:{L}{i})")
        c.number_format, c.font = HRS, BOLD
    ws.column_dimensions["A"].width = 22
    for j in range(2, len(month_tabs) + 3):
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 13

    wb.save(out_path)
    print(f"{len(month_tabs)} abas de mês · {len(agg)} linhas · {len(services)} tipos de serviço · {len(staff)} pessoas")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(*sys.argv[1:])
