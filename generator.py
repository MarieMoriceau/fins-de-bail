"""Génère le fichier de suivi commercial des fins de bail triennales.

Lit un fichier xlsx source contenant au minimum les colonnes :
- ID (col A)
- Société (col B)
- SIREN Pappers (col C)
- Siège CP (col L)
- Siège Ville (col M)
- ✅ Adresse Retenue (col AB)
- 📅 Date Adresse Retenue (col AD)

Produit un nouveau fichier xlsx avec :
- Mode d'emploi
- Fin < 6 mois (priorité HAUTE)
- Fin 6-9 mois (priorité MOYENNE)
- Fin 9-12 mois (à anticiper)
- Données (copie + colonnes calculées)
"""
from datetime import date, datetime
from io import BytesIO

import openpyxl
from dateutil.relativedelta import relativedelta
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


AUTO_COLS = [
    ("ID", 10),
    ("Société", 26),
    ("Adresse", 38),
    ("CP", 8),
    ("Ville", 18),
    ("Entrée dans les lieux", 14),
    ("Fin de bail", 13),
    ("Mois restants", 10),
    ("SIREN", 12),
]
MANUAL_COLS = [
    ("Contact", 20),
    ("Téléphone", 14),
    ("Email", 26),
    ("Date appel", 12),
    ("Statut", 14),
    ("Relance le", 12),
    ("Notes", 40),
]
CAT_META = {
    "Fin < 6 mois": ("priorité HAUTE", "C00000"),
    "Fin 6-9 mois": ("priorité MOYENNE", "ED7D31"),
    "Fin 9-12 mois": ("à anticiper", "4472C4"),
}


def _parse_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
    return None


def _months_diff(start, end):
    months = (end.year - start.year) * 12 + (end.month - start.month)
    if end.day < start.day:
        months -= 1
    return months


def _extract_rows(ws_src, today):
    seen = set()
    rows = []
    for r in range(2, ws_src.max_row + 1):
        id_ = ws_src.cell(row=r, column=1).value
        if id_ is None or id_ in seen:
            continue
        seen.add(id_)
        d = _parse_date(ws_src.cell(row=r, column=30).value)
        if d is None:
            continue
        months_since = _months_diff(d, today)
        n_cycles = months_since // 36 + 1
        end = d + relativedelta(months=n_cycles * 36)
        months_until = _months_diff(today, end)
        rows.append({
            "id": id_,
            "societe": ws_src.cell(row=r, column=2).value,
            "siren": ws_src.cell(row=r, column=3).value,
            "adresse": ws_src.cell(row=r, column=28).value,
            "cp": ws_src.cell(row=r, column=12).value,
            "ville": ws_src.cell(row=r, column=13).value,
            "entree": d,
            "fin": end,
            "mois": months_until,
        })
    return rows


def _write_call_sheet(wb, sheet_name, data_list, today):
    tag, color = CAT_META[sheet_name]
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    s = wb.create_sheet(sheet_name)
    total_cols = len(AUTO_COLS) + len(MANUAL_COLS)
    last_col = get_column_letter(total_cols)

    title_font = Font(bold=True, size=14, color="FFFFFF", name="Calibri")
    header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
    auto_fill = PatternFill("solid", start_color="1F4E78")
    manual_fill = PatternFill("solid", start_color="548235")
    border = Border(*(Side(style="thin", color="BFBFBF") for _ in range(4)))

    s.merge_cells(f"A1:{last_col}1")
    s["A1"] = f"📞 {sheet_name.upper()} — {tag} — {len(data_list)} sociétés"
    s["A1"].font = title_font
    s["A1"].fill = PatternFill("solid", start_color=color)
    s["A1"].alignment = Alignment(horizontal="center", vertical="center")
    s.row_dimensions[1].height = 30

    s.merge_cells(f"A2:{last_col}2")
    s["A2"] = (
        f"Généré le {today.strftime('%d/%m/%Y')}. "
        "Bleu = infos société (auto). Vert = à remplir après appel."
    )
    s["A2"].font = Font(italic=True, size=10, color="595959")
    s["A2"].alignment = Alignment(horizontal="center", vertical="center")
    s.row_dimensions[2].height = 22

    header_row = 3
    for i, (label, width) in enumerate(AUTO_COLS + MANUAL_COLS, 1):
        c = s.cell(row=header_row, column=i, value=label)
        c.fill = auto_fill if i <= len(AUTO_COLS) else manual_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
        s.column_dimensions[get_column_letter(i)].width = width
    s.row_dimensions[header_row].height = 32

    for idx, r in enumerate(data_list):
        excel_row = header_row + 1 + idx
        values = [r["id"], r["societe"], r["adresse"], r["cp"], r["ville"],
                  r["entree"], r["fin"], r["mois"], r["siren"]]
        for i, v in enumerate(values, 1):
            c = s.cell(row=excel_row, column=i, value=v)
            c.border = border
            c.font = Font(name="Calibri", size=10)
            c.alignment = Alignment(vertical="center")
            if i in (6, 7):
                c.number_format = "DD/MM/YYYY"
            elif i in (4, 9):
                c.number_format = "0;;;@"
        for j in range(1, len(MANUAL_COLS) + 1):
            col_i = len(AUTO_COLS) + j
            c = s.cell(row=excel_row, column=col_i)
            c.border = border
            c.fill = PatternFill("solid", start_color="EAF4E6")
            if MANUAL_COLS[j - 1][0] in ("Date appel", "Relance le"):
                c.number_format = "DD/MM/YYYY"

    extra_rows = 20
    last_data_row = header_row + len(data_list)
    for e in range(extra_rows):
        excel_row = last_data_row + 1 + e
        for i in range(1, total_cols + 1):
            c = s.cell(row=excel_row, column=i)
            c.border = border
            if i > len(AUTO_COLS):
                c.fill = PatternFill("solid", start_color="EAF4E6")

    mois_range = f"H{header_row + 1}:H{last_data_row + extra_rows}"
    rules = [
        ("lessThan", ["3"], "F8CBAD", "9C0006"),
        ("between", ["3", "5"], "FFE699", "806000"),
        ("between", ["6", "8"], "FFF2CC", "7F6000"),
        ("between", ["9", "12"], "DDEBF7", "1F4E78"),
    ]
    for op, formula, bg, fg in rules:
        s.conditional_formatting.add(
            mois_range,
            CellIsRule(operator=op, formula=formula,
                       fill=PatternFill("solid", start_color=bg),
                       font=Font(bold=True, color=fg)),
        )

    statut_letter = get_column_letter(len(AUTO_COLS) + 5)
    statut_range = f"{statut_letter}{header_row + 1}:{statut_letter}{last_data_row + extra_rows}"
    for val, bg, fg in [("OK signé", "C6EFCE", "006100"),
                         ("À rappeler", "FFEB9C", "9C5700"),
                         ("Refus", "FFC7CE", "9C0006"),
                         ("Injoignable", "D9D9D9", "595959")]:
        s.conditional_formatting.add(
            statut_range,
            CellIsRule(operator="equal", formula=[f'"{val}"'],
                       fill=PatternFill("solid", start_color=bg),
                       font=Font(bold=True, color=fg)),
        )
    dv = DataValidation(
        type="list",
        formula1='"À appeler,À rappeler,OK signé,Refus,Injoignable,Ne pas contacter"',
        allow_blank=True,
    )
    dv.add(statut_range)
    s.add_data_validation(dv)

    s.auto_filter.ref = f"A{header_row}:{last_col}{last_data_row + extra_rows}"
    s.freeze_panes = s.cell(row=header_row + 1, column=3)


def _write_readme(wb, today, counts):
    if "Mode d'emploi" in wb.sheetnames:
        del wb["Mode d'emploi"]
    r = wb.create_sheet("Mode d'emploi", 0)
    r.column_dimensions["A"].width = 28
    r.column_dimensions["B"].width = 100
    r["A1"] = "📞 Suivi commercial — fins de bail triennales"
    r.merge_cells("A1:B1")
    r["A1"].font = Font(bold=True, size=16, color="FFFFFF", name="Calibri")
    r["A1"].fill = PatternFill("solid", start_color="1F4E78")
    r["A1"].alignment = Alignment(horizontal="center", vertical="center")
    r.row_dimensions[1].height = 34

    total = sum(counts.values())
    lines = [
        ("", ""),
        ("Date de calcul", f"{today.strftime('%d/%m/%Y')} — {total} sociétés au total (doublons supprimés)."),
        ("", ""),
        ("🔴 Fin < 6 mois", f"{counts['Fin < 6 mois']} sociétés — priorité HAUTE."),
        ("🟠 Fin 6-9 mois", f"{counts['Fin 6-9 mois']} sociétés — priorité MOYENNE."),
        ("🔵 Fin 9-12 mois", f"{counts['Fin 9-12 mois']} sociétés — à anticiper."),
        ("", ""),
        ("Colonnes bleues", "Auto (ID, Société, Adresse, CP, Ville, Entrée, Fin de bail, Mois restants, SIREN)."),
        ("Colonnes vertes", "À remplir (Contact, Téléphone, Email, Date appel, Statut, Relance, Notes)."),
        ("Statut", "Liste déroulante : À appeler, À rappeler, OK signé, Refus, Injoignable, Ne pas contacter."),
        ("Code couleur Mois", "Rouge < 3 • Orange 3-5 • Jaune 6-8 • Bleu 9-12."),
    ]
    for i, (k, v) in enumerate(lines, 2):
        r.cell(row=i, column=1, value=k).font = Font(bold=True, name="Calibri", size=11)
        r.cell(row=i, column=1).alignment = Alignment(vertical="top", wrap_text=True)
        r.cell(row=i, column=2, value=v).alignment = Alignment(wrap_text=True, vertical="top")
        r.cell(row=i, column=2).font = Font(name="Calibri", size=11)
        r.row_dimensions[i].height = 26


def generate(source_bytes: bytes, today=None) -> bytes:
    """Prend un xlsx source en bytes, renvoie le xlsx généré en bytes."""
    today = today or date.today()
    wb_src = openpyxl.load_workbook(BytesIO(source_bytes), data_only=True)
    ws_src = wb_src.active

    rows = _extract_rows(ws_src, today)
    cat_rows = {
        "Fin < 6 mois":  sorted([r for r in rows if r["mois"] < 6], key=lambda x: (x["mois"], x["fin"])),
        "Fin 6-9 mois":  sorted([r for r in rows if 6 <= r["mois"] < 9], key=lambda x: (x["mois"], x["fin"])),
        "Fin 9-12 mois": sorted([r for r in rows if 9 <= r["mois"] <= 12], key=lambda x: (x["mois"], x["fin"])),
    }
    counts = {k: len(v) for k, v in cat_rows.items()}

    # Workbook de sortie : on repart du fichier source en ajoutant les onglets
    wb = openpyxl.load_workbook(BytesIO(source_bytes))
    # Renomme l'onglet principal en "Données" si besoin
    main = wb.active
    if main.title != "Données":
        main.title = "Données"

    for sheet_name, data_list in cat_rows.items():
        _write_call_sheet(wb, sheet_name, data_list, today)

    _write_readme(wb, today, counts)

    order = ["Mode d'emploi", "Fin < 6 mois", "Fin 6-9 mois", "Fin 9-12 mois", "Données"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames]

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
