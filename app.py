import io
import math
import json
import os
from datetime import datetime
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
CORS(app)  # autorise les appels depuis le navigateur

# ─────────────────────────────────────────────
# Générateur Excel (identique à ton fichier original)
# ─────────────────────────────────────────────
def generate_constat_xlsx(data: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Constat"

    col_widths = {
        'A': 3, 'B': 22, 'C': 6, 'D': 10, 'E': 4, 'F': 10,
        'G': 4, 'H': 14, 'I': 6, 'J': 6, 'K': 4,
        'L': 10, 'M': 6, 'N': 6, 'O': 4, 'P': 10, 'Q': 4, 'R': 4
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    def thin():
        s = Side(style='thin')
        return Border(left=s, right=s, top=s, bottom=s)

    def cell(row, col, value=None, bold=False, size=10, color="000000",
             bg=None, align="left", wrap=False, italic=False, number_format=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(name="Arial", bold=bold, size=size, color=color, italic=italic)
        c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
        if bg:
            c.fill = PatternFill("solid", fgColor=bg)
        if number_format:
            c.number_format = number_format
        return c

    BLUE_DARK  = "1F3864"
    BLUE_MED   = "2E75B6"
    BLUE_LIGHT = "BDD7EE"
    GRAY_LIGHT = "F2F2F2"
    GREEN_OK   = "E2EFDA";  GREEN_TXT = "375623"
    RED_NOK    = "FCE4D6";  RED_TXT   = "843C0C"
    YELLOW_HDR = "FFF2CC";  WHITE     = "FFFFFF"

    ws.row_dimensions[1].height  = 30
    ws.row_dimensions[19].height = 20
    ws.row_dimensions[20].height = 20
    ws.row_dimensions[22].height = 20

    site = data.get('site', 'Reichshoffen')
    norme = data.get('norme_visuelle', 'DTREI 150613')

    # ROW 1
    ws.merge_cells('B1:G1')
    c = ws['B1']
    c.value = "VERIFICATION\nETALON SECONDAIRE DE PEINTURE"
    c.font = Font(name="Arial", bold=True, size=12, color=WHITE)
    c.fill = PatternFill("solid", fgColor=BLUE_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells('H1:J1')
    c = ws['H1']
    c.value = "Constat N°"
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    c.fill = PatternFill("solid", fgColor=BLUE_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells('K1:R1')
    c = ws['K1']
    c.value = f"Site de {site}"
    c.font = Font(name="Arial", bold=True, size=10, color=WHITE)
    c.fill = PatternFill("solid", fgColor=BLUE_DARK)
    c.alignment = Alignment(horizontal="right", vertical="center")

    # ROW 3
    ws.merge_cells('H3:R3')
    c = ws['H3']
    c.value = data['numero_constat']
    c.font = Font(name="Arial", bold=True, size=10, color=BLUE_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=BLUE_LIGHT)

    # ROWS 5-7
    ws.merge_cells('B5:F5')
    cell(5, 2, "Référence de l'étalon :", bold=True, bg=GRAY_LIGHT)
    ws.merge_cells('H5:R5')
    c = ws['H5']
    c.value = f"Teinte : {data['teinte']}"
    c.font = Font(name="Arial", bold=True, size=10, color=BLUE_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells('H6:R6')
    ws['H6'].value = f"Brillance : {data['brillance']}"
    ws['H6'].font = Font(name="Arial", size=10)

    ws.merge_cells('H7:R7')
    ws['H7'].value = f"Technologie : {data['technologie']}"
    ws['H7'].font = Font(name="Arial", size=10)

    # ROW 9-11 Spectro
    cell(9, 2, "Référence spectro-colorimètre", bold=True, size=9, bg=GRAY_LIGHT)
    ws.merge_cells('E9:G9')
    cell(9, 5, data['spectro_ref'], size=9, bg=YELLOW_HDR, bold=True)
    ws.merge_cells('H9:R9')
    cell(9, 8, "N° ETALON SECONDAIRE A VALIDER", bold=True, size=9, bg=BLUE_LIGHT, align="center")

    cell(10, 2, "N° d'identification", size=9)
    ws.merge_cells('E10:G10')
    cell(10, 5, data['spectro_id'], size=9, bg=YELLOW_HDR)
    ws.merge_cells('H10:R10')
    cell(10, 8, data['etalon_secondaire'], bold=True, size=11, color=BLUE_DARK, align="center")

    cell(11, 2, "Validité", size=9)
    ws.merge_cells('E11:G11')
    c = ws.cell(row=11, column=5, value=data['spectro_validite'])
    c.font = Font(name="Arial", size=9, color="CC0000")
    c.number_format = "DD/MM/YYYY"
    ws.merge_cells('H11:R11')
    cell(11, 8, "Valeur L * a * b", bold=True, size=9, bg=BLUE_LIGHT, align="center")

    # ROW 12-13 Primaire
    cell(12, 2, "N° ETALON PRIMAIRE DE REFERENCE", bold=True, size=9, bg=GRAY_LIGHT)
    cell(12, 8, "L", bold=True, align="center", bg=BLUE_LIGHT, size=9)
    cell(12, 12, "a", bold=True, align="center", bg=BLUE_LIGHT, size=9)
    cell(12, 16, "b", bold=True, align="center", bg=BLUE_LIGHT, size=9)

    ws.merge_cells('B13:G13')
    cell(13, 2, data['etalon_primaire'], bold=True, size=10, color=BLUE_DARK)
    ws.merge_cells('H13:K13')
    cell(13, 8, data['ref_L'], align="center", size=10, bg=YELLOW_HDR)
    ws.merge_cells('L13:O13')
    cell(13, 12, data['ref_a'], align="center", size=10, bg=YELLOW_HDR)
    ws.merge_cells('P13:R13')
    cell(13, 16, data['ref_b'], align="center", size=10, bg=YELLOW_HDR)

    # ROW 14-15 Deltas headers + tolerances
    cell(14, 2, "Date de validation étalon primaire", size=9, bg=GRAY_LIGHT)
    ws.merge_cells('E14:G14')
    c = ws.cell(row=14, column=5, value=data['date_validation_primaire'])
    c.font = Font(name="Arial", size=9)
    c.number_format = "DD/MM/YYYY"
    cell(14, 8,  "∆L", bold=True, align="center", bg=BLUE_LIGHT, size=9)
    cell(14, 12, "∆a", bold=True, align="center", bg=BLUE_LIGHT, size=9)
    cell(14, 16, "∆b", bold=True, align="center", bg=BLUE_LIGHT, size=9)

    cell(15, 8,  data['tol_L_min'], align="center", size=9, italic=True, color="666666")
    cell(15, 9,  data['tol_L_max'], align="center", size=9, italic=True, color="666666")
    cell(15, 12, data['tol_a_min'], align="center", size=9, italic=True, color="666666")
    cell(15, 13, data['tol_a_max'], align="center", size=9, italic=True, color="666666")
    cell(15, 16, data['tol_b_min'], align="center", size=9, italic=True, color="666666")
    cell(15, 17, data['tol_b_max'], align="center", size=9, italic=True, color="666666")

    # ROW 16 — Deltas calculés (formules Excel)
    ws.merge_cells('H16:K16')
    ws['H16'] = '=B19-H13'
    ws['H16'].number_format = '0.00'
    ws['H16'].font = Font(name="Arial", size=10, bold=True)
    ws['H16'].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells('L16:O16')
    ws['L16'] = '=D19-L13'
    ws['L16'].number_format = '0.00'
    ws['L16'].font = Font(name="Arial", size=10, bold=True)
    ws['L16'].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells('P16:R16')
    ws['P16'] = '=F19-P13'
    ws['P16'].number_format = '0.00'
    ws['P16'].font = Font(name="Arial", size=10, bold=True)
    ws['P16'].alignment = Alignment(horizontal="center", vertical="center")

    # ROW 17-18 — Headers mesures
    ws.merge_cells('B17:C17')
    cell(17, 2, "Valeurs mesurées", bold=True, size=9, bg=GRAY_LIGHT)
    cell(17, 4, "L",  bold=True, align="center", bg=BLUE_MED, color=WHITE, size=9)
    cell(17, 6, "a",  bold=True, align="center", bg=BLUE_MED, color=WHITE, size=9)
    cell(17, 8, "b",  bold=True, align="center", bg=BLUE_MED, color=WHITE, size=9)
    cell(17, 10, "∆E", bold=True, align="center", bg=BLUE_MED, color=WHITE, size=9)
    cell(18, 10, f"≤ {data['tol_dE']}", align="center", size=9, italic=True, color="666666")

    # ROW 19 — Valeurs mesurées
    for col, val in [(2, data['mes_L']), (4, data['mes_a']), (6, data['mes_b'])]:
        c = ws.cell(row=19, column=col, value=val)
        c.font = Font(name="Arial", size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", fgColor=YELLOW_HDR)
        c.number_format = '0.00'

    ws['J19'] = '=SQRT((B19-H13)^2+(D19-L13)^2+(F19-P13)^2)'
    ws['J19'].number_format = '0.00'
    ws['J19'].font = Font(name="Arial", size=11, bold=True)
    ws['J19'].alignment = Alignment(horizontal="center", vertical="center")
    ws['J19'].fill = PatternFill("solid", fgColor="E8F0FE")

    # ROW 20 — Résultat colorimétrique
    ws.merge_cells('B20:G20')
    cell(20, 2, "Résultat du contrôle colorimétrique :", bold=True, size=10)

    dL = data['mes_L'] - data['ref_L']
    da = data['mes_a'] - data['ref_a']
    db = data['mes_b'] - data['ref_b']
    dE = math.sqrt(dL**2 + da**2 + db**2)
    ok = (data['tol_L_min'] <= dL <= data['tol_L_max'] and
          data['tol_a_min'] <= da <= data['tol_a_max'] and
          data['tol_b_min'] <= db <= data['tol_b_max'] and
          dE <= data['tol_dE'])

    ws.merge_cells('H20:R20')
    c = ws['H20']
    c.value = "CONFORME" if ok else "NON CONFORME"
    c.font = Font(name="Arial", bold=True, size=12, color=GREEN_TXT if ok else RED_TXT)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=GREEN_OK if ok else RED_NOK)

    # ROW 22 — Résultat visuel
    ws.merge_cells('B22:G22')
    cell(22, 2, f"Résultat du contrôle visuel suivant {norme} :", bold=True, size=10)
    rv = data.get('res_visuel', 'CONFORME')
    ws.merge_cells('H22:R22')
    c = ws['H22']
    c.value = rv
    c.font = Font(name="Arial", bold=True, size=12, color=GREEN_TXT if rv == "CONFORME" else RED_TXT)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=GREEN_OK if rv == "CONFORME" else RED_NOK)

    # ROW 24 — Signataire
    cell(24, 2, "Etabli par", bold=True, size=9, bg=GRAY_LIGHT)
    ws.merge_cells('E24:G24')
    cell(24, 5, data['etabli_par'], size=10, bold=True)
    cell(24, 8, "Le", bold=True, size=9)
    ws.merge_cells('I24:M24')
    c = ws.cell(row=24, column=9, value=data['date_constat'])
    c.font = Font(name="Arial", size=10, bold=True)
    c.number_format = "DD/MM/YYYY"
    c.alignment = Alignment(horizontal="left", vertical="center")

    ws.print_area = 'A1:R25'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = 'landscape'

    # Retourne les bytes du fichier (pas de fichier sur disque)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# Routes Flask
# ─────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "message": "Constat Generator API is running"})


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    if not data:
        return jsonify({"error": "JSON body requis"}), 400

    required = ["numero_constat", "teinte", "brillance", "technologie",
                "spectro_ref", "spectro_id", "spectro_validite",
                "etalon_secondaire", "etalon_primaire",
                "ref_L", "ref_a", "ref_b",
                "date_validation_primaire",
                "tol_L_min", "tol_L_max",
                "tol_a_min", "tol_a_max",
                "tol_b_min", "tol_b_max", "tol_dE",
                "mes_L", "mes_a", "mes_b",
                "etabli_par", "date_constat"]

    missing = [f for f in required if f not in data]
    if missing:
        return jsonify({"error": f"Champs manquants : {missing}"}), 400

    try:
        xlsx_bytes = generate_constat_xlsx(data)
        filename = data['numero_constat'].replace("/", "_").replace("°", "") + ".xlsx"
        return send_file(
            io.BytesIO(xlsx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("🎨 Constat Generator API — http://localhost:5000")
    print("   POST /generate  →  retourne le .xlsx")
    print("   GET  /health    →  vérification")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
