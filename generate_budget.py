"""Genererar en månadsbudget-arbetsbok i SEK (Monthly_Budget_SEK.xlsx).

Kör:  python3 generate_budget.py
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation


OUTPUT = Path(__file__).with_name("Monthly_Budget_SEK.xlsx")

SEK_FMT = '#,##0" kr";[Red]-#,##0" kr";0" kr"'
PCT_FMT = "0.0%"
DATE_FMT = "yyyy-mm-dd"

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SUBHEAD_FILL = PatternFill("solid", fgColor="2E75B6")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
ACCENT_FILL = PatternFill("solid", fgColor="FFF2CC")

WHITE_BOLD = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
BOLD = Font(name="Calibri", size=11, bold=True)
TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="1F4E78")
KPI_LABEL_FONT = Font(name="Calibri", size=11, bold=True, color="404040")
KPI_VALUE_FONT = Font(name="Calibri", size=14, bold=True, color="1F4E78")

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")

CATEGORIES = [
    "Boende",
    "Mat",
    "Transport",
    "Nöje",
    "Räkningar",
    "Försäkring",
    "Prenumerationer",
    "Hälsa",
    "Kläder",
    "Övrigt",
]

EXPENSE_ROWS = 60  # antal förformaterade rader på Utgifter
SAVINGS_ROWS = 8


def style_headers(ws, row: int, last_col: int) -> None:
    for c in range(1, last_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = WHITE_BOLD
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER
    ws.row_dimensions[row].height = 22


def set_widths(ws, widths: dict[str, float]) -> None:
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# --- Inställningar ----------------------------------------------------------

def build_settings(wb: Workbook):
    ws = wb.create_sheet("Inställningar")
    ws["A1"] = "Kategorier"
    ws["A1"].font = WHITE_BOLD
    ws["A1"].fill = HEADER_FILL
    ws["A1"].alignment = CENTER
    for i, cat in enumerate(CATEGORIES, start=2):
        ws.cell(row=i, column=1, value=cat).alignment = LEFT

    ws["C1"] = "Tips"
    ws["C1"].font = BOLD
    ws["C2"] = "Ändra eller lägg till kategorier ovan."
    ws["C3"] = "Dropdown-menyn på fliken Utgifter uppdateras automatiskt"
    ws["C4"] = "om du utvidgar listan inom A2:A50."

    set_widths(ws, {"A": 22, "C": 60})
    ws.sheet_view.showGridLines = False
    return ws


# --- Inkomst ----------------------------------------------------------------

INCOME_DEFAULTS = [
    ("Lön efter skatt", 27840, 27840),
    ("Övrig inkomst", 0, 0),
]


def build_income(wb: Workbook):
    ws = wb.create_sheet("Inkomst")
    headers = ["Källa", "Budget (kr)", "Faktiskt (kr)", "Differens (kr)"]
    ws.append(headers)
    style_headers(ws, 1, len(headers))

    start = 2
    for i, (name, budget, actual) in enumerate(INCOME_DEFAULTS):
        r = start + i
        ws.cell(row=r, column=1, value=name).alignment = LEFT
        ws.cell(row=r, column=2, value=budget).number_format = SEK_FMT
        ws.cell(row=r, column=3, value=actual).number_format = SEK_FMT
        ws.cell(row=r, column=4, value=f"=C{r}-B{r}").number_format = SEK_FMT
        for c in range(1, 5):
            ws.cell(row=r, column=c).border = BORDER

    last_row = start + len(INCOME_DEFAULTS) - 1
    total_row = last_row + 1
    ws.cell(row=total_row, column=1, value="TOTALT").font = BOLD
    ws.cell(row=total_row, column=2, value=f"=SUM(B{start}:B{last_row})")
    ws.cell(row=total_row, column=3, value=f"=SUM(C{start}:C{last_row})")
    ws.cell(row=total_row, column=4, value=f"=SUM(D{start}:D{last_row})")
    for c in range(1, 5):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = BOLD
        cell.border = BORDER
        if c >= 2:
            cell.number_format = SEK_FMT

    diff_range = f"D{start}:D{last_row}"
    ws.conditional_formatting.add(
        diff_range,
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="F8CBAD")),
    )
    ws.conditional_formatting.add(
        diff_range,
        CellIsRule(operator="greaterThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="C6E0B4")),
    )

    set_widths(ws, {"A": 26, "B": 16, "C": 16, "D": 18})
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws, total_row


# --- Utgifter ---------------------------------------------------------------

def build_expenses(wb: Workbook):
    ws = wb.create_sheet("Utgifter")
    headers = ["Datum", "Kategori", "Beskrivning",
               "Budget (kr)", "Faktiskt (kr)", "Differens (kr)", "Status"]
    ws.append(headers)
    style_headers(ws, 1, len(headers))

    start = 2
    end = start + EXPENSE_ROWS - 1

    # Återkommande månadsposter – datumkolumnen lämnas tom så du kan
    # fylla i när varje räkning faktiskt betalas.
    sample = [
        (None, "Boende",          "Hyra",                                 3914, 0),
        (None, "Mat",             "Matbudget",                            6200, 0),
        (None, "Transport",       "Diesel",                               2750, 0),
        (None, "Transport",       "Underhåll bil/mc",                     1000, 0),
        (None, "Transport",       "Bilskatt (5639 kr/år ÷ 12)",            470, 0),
        (None, "Försäkring",      "Bilförsäkring (416 kr/år ÷ 12)",         35, 0),
        (None, "Försäkring",      "A-kassa",                               160, 0),
        (None, "Prenumerationer", "Spotify",                               129, 0),
        (None, "Prenumerationer", "Apple",                                  39, 0),
        (None, "Prenumerationer", "Google",                                 40, 0),
        (None, "Prenumerationer", "Soundcloud",                             75, 0),
        (None, "Prenumerationer", "Claude Max",                           1300, 0),
        (None, "Prenumerationer", "Systeme.io",                            160, 0),
        (None, "Hälsa",           "Linser",                                500, 0),
        (None, "Kläder",          "Kläder och övriga utgifter",              0, 0),
    ]

    for i in range(EXPENSE_ROWS):
        r = start + i
        if i < len(sample):
            d, cat, desc, b, a = sample[i]
            ws.cell(row=r, column=1, value=d)
            ws.cell(row=r, column=2, value=cat)
            ws.cell(row=r, column=3, value=desc)
            ws.cell(row=r, column=4, value=b)
            ws.cell(row=r, column=5, value=a)
        ws.cell(row=r, column=1).number_format = DATE_FMT
        ws.cell(row=r, column=4).number_format = SEK_FMT
        ws.cell(row=r, column=5).number_format = SEK_FMT
        ws.cell(row=r, column=6,
                value=f'=IF(AND(D{r}=0,E{r}=0),"",D{r}-E{r})')
        ws.cell(row=r, column=6).number_format = SEK_FMT
        ws.cell(row=r, column=7,
                value=(f'=IF(AND(D{r}=0,E{r}=0),"",'
                       f'IF(E{r}>D{r},"Över","OK"))'))
        ws.cell(row=r, column=7).alignment = CENTER
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = BORDER

    total_row = end + 1
    ws.cell(row=total_row, column=1, value="TOTALT").font = BOLD
    ws.cell(row=total_row, column=4, value=f"=SUM(D{start}:D{end})")
    ws.cell(row=total_row, column=5, value=f"=SUM(E{start}:E{end})")
    ws.cell(row=total_row, column=6, value=f"=SUM(D{start}:D{end})-SUM(E{start}:E{end})")
    for c in range(1, 8):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = BOLD
        cell.border = BORDER
        if c in (4, 5, 6):
            cell.number_format = SEK_FMT

    # Dropdown-validering på Kategori (B-kolumnen)
    dv = DataValidation(type="list", formula1="=Kategorier", allow_blank=True,
                        showErrorMessage=True,
                        errorTitle="Okänd kategori",
                        error="Välj en kategori från listan på Inställningar.")
    dv.add(f"B{start}:B{end}")
    ws.add_data_validation(dv)

    # Differens: röd om <0, grön om >0
    diff_range = f"F{start}:F{end}"
    ws.conditional_formatting.add(
        diff_range,
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="F8CBAD")),
    )
    ws.conditional_formatting.add(
        diff_range,
        CellIsRule(operator="greaterThan", formula=["0"],
                   fill=PatternFill("solid", fgColor="C6E0B4")),
    )
    # Status: röd "Över", grön "OK"
    status_range = f"G{start}:G{end}"
    ws.conditional_formatting.add(
        status_range,
        FormulaRule(formula=[f'$G{start}="Över"'],
                    fill=PatternFill("solid", fgColor="F4B084"),
                    font=Font(bold=True, color="9C0006")),
    )
    ws.conditional_formatting.add(
        status_range,
        FormulaRule(formula=[f'$G{start}="OK"'],
                    fill=PatternFill("solid", fgColor="C6EFCE"),
                    font=Font(bold=True, color="006100")),
    )

    # Sidopanel: per kategori-summor (driver dashboard-diagrammen)
    ws.cell(row=1, column=9, value="Kategori").font = WHITE_BOLD
    ws.cell(row=1, column=9).fill = SUBHEAD_FILL
    ws.cell(row=1, column=9).alignment = CENTER
    ws.cell(row=1, column=10, value="Budget").font = WHITE_BOLD
    ws.cell(row=1, column=10).fill = SUBHEAD_FILL
    ws.cell(row=1, column=10).alignment = CENTER
    ws.cell(row=1, column=11, value="Faktiskt").font = WHITE_BOLD
    ws.cell(row=1, column=11).fill = SUBHEAD_FILL
    ws.cell(row=1, column=11).alignment = CENTER

    for i, cat in enumerate(CATEGORIES, start=2):
        ws.cell(row=i, column=9, value=cat).alignment = LEFT
        ws.cell(row=i, column=10,
                value=f'=SUMIF(B{start}:B{end},I{i},D{start}:D{end})'
                ).number_format = SEK_FMT
        ws.cell(row=i, column=11,
                value=f'=SUMIF(B{start}:B{end},I{i},E{start}:E{end})'
                ).number_format = SEK_FMT
        for c in range(9, 12):
            ws.cell(row=i, column=c).border = BORDER

    set_widths(ws, {
        "A": 12, "B": 18, "C": 28, "D": 14, "E": 14,
        "F": 16, "G": 10, "H": 2, "I": 18, "J": 14, "K": 14,
    })
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws, start, end, total_row


# --- Konton -----------------------------------------------------------------

ACCOUNT_DEFAULTS = [
    ("Sparkonto", 34000, "Behöver tas pengar härifrån denna månad"),
    ("Allkonto",   2600, ""),
]


def build_accounts(wb: Workbook):
    ws = wb.create_sheet("Konton")
    headers = ["Konto", "Saldo (kr)", "Anteckning"]
    ws.append(headers)
    style_headers(ws, 1, len(headers))

    start = 2
    for i, (name, balance, note) in enumerate(ACCOUNT_DEFAULTS):
        r = start + i
        ws.cell(row=r, column=1, value=name).alignment = LEFT
        ws.cell(row=r, column=2, value=balance).number_format = SEK_FMT
        ws.cell(row=r, column=3, value=note).alignment = LEFT
        for c in range(1, 4):
            ws.cell(row=r, column=c).border = BORDER

    last = start + len(ACCOUNT_DEFAULTS) - 1
    total_row = last + 1
    ws.cell(row=total_row, column=1, value="TOTALT").font = BOLD
    ws.cell(row=total_row, column=2, value=f"=SUM(B{start}:B{last})")
    for c in range(1, 4):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = BOLD
        cell.border = BORDER
    ws.cell(row=total_row, column=2).number_format = SEK_FMT

    set_widths(ws, {"A": 22, "B": 16, "C": 50})
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws, total_row


# --- Sparmål ----------------------------------------------------------------

SAVINGS_DEFAULTS = [
    ("Buffert",   60000, 34000, 0),
    ("Snöskoter", 40000,     0, 0),
]


def build_savings(wb: Workbook):
    ws = wb.create_sheet("Sparmål")
    headers = [
        "Mål", "Målbelopp (kr)", "Sparat hittills (kr)",
        "Månadsinsättning (kr)", "Återstår (kr)",
        "Månader kvar", "Måldatum", "Framsteg %",
    ]
    ws.append(headers)
    style_headers(ws, 1, len(headers))

    start = 2
    end = start + SAVINGS_ROWS - 1

    for i in range(SAVINGS_ROWS):
        r = start + i
        if i < len(SAVINGS_DEFAULTS):
            name, target, saved, monthly = SAVINGS_DEFAULTS[i]
            ws.cell(row=r, column=1, value=name).alignment = LEFT
            ws.cell(row=r, column=2, value=target)
            ws.cell(row=r, column=3, value=saved)
            ws.cell(row=r, column=4, value=monthly)
        for c in (2, 3, 4, 5):
            ws.cell(row=r, column=c).number_format = SEK_FMT
        ws.cell(row=r, column=5,
                value=f'=IF(B{r}="","",MAX(0,B{r}-C{r}))')
        ws.cell(row=r, column=6,
                value=f'=IFERROR(IF(D{r}>0,CEILING(E{r}/D{r},1),""),"")')
        ws.cell(row=r, column=6).alignment = CENTER
        ws.cell(row=r, column=7,
                value=f'=IFERROR(EDATE(TODAY(),F{r}),"")')
        ws.cell(row=r, column=7).number_format = DATE_FMT
        ws.cell(row=r, column=7).alignment = CENTER
        ws.cell(row=r, column=8,
                value=f'=IFERROR(MIN(1,C{r}/B{r}),0)')
        ws.cell(row=r, column=8).number_format = PCT_FMT
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = BORDER

    # Datafält (data bar) på Framsteg %
    ws.conditional_formatting.add(
        f"H{start}:H{end}",
        DataBarRule(start_type="num", start_value=0,
                    end_type="num", end_value=1,
                    color="63BE7B", showValue=True),
    )

    set_widths(ws, {
        "A": 28, "B": 16, "C": 18, "D": 20, "E": 16,
        "F": 14, "G": 14, "H": 14,
    })
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    return ws, start, end


# --- Översikt (Dashboard) ---------------------------------------------------

def build_dashboard(wb: Workbook, income_total_row: int,
                    expense_start: int, expense_end: int,
                    savings_start: int, savings_end: int,
                    accounts_total_row: int):
    ws = wb.create_sheet("Översikt", 0)  # första fliken

    ws.merge_cells("A1:F1")
    ws["A1"] = "Månadsbudget – SEK"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 32

    ws["A3"] = "Period"
    ws["A3"].font = KPI_LABEL_FONT
    ws["B3"] = "=TEXT(TODAY(),\"yyyy-mm\")"
    ws["B3"].font = KPI_VALUE_FONT
    ws["B3"].alignment = LEFT

    kpi = [
        ("Total inkomst",  f"=Inkomst!C{income_total_row}",                SEK_FMT),
        ("Total utgift (budget)",
                           f"=SUM(Utgifter!D{expense_start}:D{expense_end})", SEK_FMT),
        ("Netto enligt budget", "=B5-B6",                                  SEK_FMT),
        ("Sparkvot",       '=IFERROR(B7/B5,0)',                            PCT_FMT),
        ("Faktiskt utgift hittills",
                           f"=SUM(Utgifter!E{expense_start}:E{expense_end})", SEK_FMT),
        ("Saldo på konton",
                           f"=Konton!B{accounts_total_row}",               SEK_FMT),
    ]
    for i, (label, formula, fmt) in enumerate(kpi):
        r = 5 + i
        ws.cell(row=r, column=1, value=label).font = KPI_LABEL_FONT
        ws.cell(row=r, column=1).alignment = LEFT
        cell = ws.cell(row=r, column=2, value=formula)
        cell.font = KPI_VALUE_FONT
        cell.number_format = fmt
        cell.alignment = LEFT
        cell.fill = ACCENT_FILL
        cell.border = BORDER

    ws["A12"] = "Sparmål – framsteg"
    ws["A12"].font = BOLD
    ws.merge_cells("A12:F12")

    headers = ["Mål", "Sparat", "Mål", "Framsteg %", "Månader kvar", "Måldatum"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=13, column=c, value=h)
        cell.font = WHITE_BOLD
        cell.fill = SUBHEAD_FILL
        cell.alignment = CENTER
        cell.border = BORDER

    for i in range(savings_end - savings_start + 1):
        src = savings_start + i
        dst = 14 + i
        ws.cell(row=dst, column=1, value=f"=Sparmål!A{src}").alignment = LEFT
        ws.cell(row=dst, column=2, value=f"=Sparmål!C{src}").number_format = SEK_FMT
        ws.cell(row=dst, column=3, value=f"=Sparmål!B{src}").number_format = SEK_FMT
        ws.cell(row=dst, column=4, value=f"=Sparmål!H{src}").number_format = PCT_FMT
        ws.cell(row=dst, column=5, value=f"=Sparmål!F{src}").alignment = CENTER
        ws.cell(row=dst, column=6, value=f"=Sparmål!G{src}").number_format = DATE_FMT
        for c in range(1, 7):
            ws.cell(row=dst, column=c).border = BORDER

    ws.conditional_formatting.add(
        f"D14:D{14 + (savings_end - savings_start)}",
        DataBarRule(start_type="num", start_value=0,
                    end_type="num", end_value=1,
                    color="63BE7B", showValue=True),
    )

    # Pie chart: faktisk fördelning per kategori
    pie = PieChart()
    pie.title = "Utgiftsfördelning (Faktiskt)"
    labels = Reference(wb["Utgifter"], min_col=9, min_row=2,
                       max_row=1 + len(CATEGORIES))
    data = Reference(wb["Utgifter"], min_col=11, min_row=1,
                     max_row=1 + len(CATEGORIES))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.height = 9
    pie.width = 14
    ws.add_chart(pie, "H3")

    # Bar chart: Budget vs Faktiskt per kategori
    bar = BarChart()
    bar.type = "bar"
    bar.style = 11
    bar.title = "Budget vs Faktiskt"
    bar.y_axis.title = "Kategori"
    bar.x_axis.title = "kr"
    cats = Reference(wb["Utgifter"], min_col=9, min_row=2,
                     max_row=1 + len(CATEGORIES))
    vals = Reference(wb["Utgifter"], min_col=10, max_col=11,
                     min_row=1, max_row=1 + len(CATEGORIES))
    bar.add_data(vals, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 10
    bar.width = 18
    ws.add_chart(bar, "H22")

    set_widths(ws, {
        "A": 22, "B": 22, "C": 18, "D": 16,
        "E": 16, "F": 16, "G": 2, "H": 12,
    })
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    return ws


# --- Bygg arbetsboken -------------------------------------------------------

def build() -> None:
    wb = Workbook()
    # Ta bort default-bladet
    wb.remove(wb.active)

    settings = build_settings(wb)
    income_ws, income_total_row = build_income(wb)
    expense_ws, exp_start, exp_end, _ = build_expenses(wb)
    savings_ws, sav_start, sav_end = build_savings(wb)
    accounts_ws, accounts_total_row = build_accounts(wb)
    dashboard = build_dashboard(
        wb, income_total_row, exp_start, exp_end,
        sav_start, sav_end, accounts_total_row,
    )

    # Namngivna områden
    last_cat_row = 1 + len(CATEGORIES)
    wb.defined_names["Kategorier"] = DefinedName(
        "Kategorier",
        attr_text=f"Inställningar!$A$2:$A${last_cat_row}",
    )
    wb.defined_names["TotalInkomst"] = DefinedName(
        "TotalInkomst", attr_text="Översikt!$B$5",
    )
    wb.defined_names["TotalUtgift"] = DefinedName(
        "TotalUtgift", attr_text="Översikt!$B$6",
    )
    wb.defined_names["Netto"] = DefinedName(
        "Netto", attr_text="Översikt!$B$7",
    )

    # Översikt först
    wb.active = wb.sheetnames.index("Översikt")

    wb.save(OUTPUT)
    print(f"Skapade {OUTPUT}")


if __name__ == "__main__":
    build()
