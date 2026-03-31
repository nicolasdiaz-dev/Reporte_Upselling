#!/usr/bin/env python3
"""
Generador de Reporte Mensual - Upselling Argentina
GOUT Global Outsourcing
"""

import io
import gspread
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from google.oauth2.service_account import Credentials
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.utils import ImageReader

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACION  (editar cada mes)
# ─────────────────────────────────────────────────────────────────────────────
CREDENTIALS_FILE = "credentials.json"
SHEET1_ID = "1umPdIIJ3v3CBBSeNKKqylZ4t2EKLtm0Ne0UBVR_gmRo"
SHEET2_ID = "1Lq6QeGJJvGM6XbdDATeps6CvFr4mye6bk6a_wtuiWuY"

MES    = "Marzo"
ANO    = "2026"
OUTPUT = f"Reporte_Upselling_{MES}{ANO}.pdf"

AUTOR_NOMBRE   = "Nicolas Diaz"
AUTOR_TEL      = "+54 11 2625-1198"
AUTOR_LINKEDIN = "linkedin.com/in/nicolas-diaz-641a17346"
AUTOR_GITHUB   = "github.com/nicolasdiaz-dev"

# ─────────────────────────────────────────────────────────────────────────────
# COLORES
# ─────────────────────────────────────────────────────────────────────────────
W, H = A4

C_DARK  = HexColor('#1D3D2F')
C_MED   = HexColor('#2D5A3D')
C_SEC   = HexColor('#4A7C6F')
C_LIGHT = HexColor('#C5E8C5')
C_PALE  = HexColor('#EAF5EA')
C_PINK  = HexColor('#F8BBD9')
C_TEAL  = HexColor('#80CBC4')
C_GRAY  = HexColor('#9E9E9E')
C_WHITE = white

M_DARK  = '#1D3D2F'
M_MED   = '#2D5A3D'
M_TEAL  = '#80CBC4'
M_LIGHT = '#C5E8C5'

# Anclas para navegacion interna
ANCHOR_RESUMEN    = "sec_resumen"
ANCHOR_ADICIONAL  = "sec_adicional"
ANCHOR_DETALLE    = "sec_detalle"
ANCHOR_ASISTENCIA = "sec_asistencia"
ANCHOR_RECOMEND   = "sec_recomend"


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def parse_int(v):
    try:
        return int(str(v).strip())
    except:
        return 0

def parse_pct(v):
    try:
        return float(str(v).replace(',', '.').replace('%', '').strip())
    except:
        return 0.0

def clean_name(raw):
    return raw.replace('_UP', '').replace('_', ' ').title().strip()

def fig_to_img(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=220, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    return ImageReader(buf)

def rr(c, x, y, w, h, r=8, fill=None, stroke=None, lw=1):
    """Rectangulo con esquinas redondeadas."""
    p = c.beginPath()
    p.moveTo(x + r, y)
    p.lineTo(x + w - r, y);          p.arcTo(x+w-2*r, y,       x+w,   y+2*r,   -90, 90)
    p.lineTo(x + w,     y + h - r);  p.arcTo(x+w-2*r, y+h-2*r, x+w,   y+h,       0, 90)
    p.lineTo(x + r,     y + h);      p.arcTo(x,       y+h-2*r, x+2*r, y+h,      90, 90)
    p.lineTo(x,         y + r);      p.arcTo(x,       y,       x+2*r, y+2*r,   180, 90)
    p.close()
    if fill:   c.setFillColor(fill)
    if stroke: c.setStrokeColor(stroke); c.setLineWidth(lw)
    c.drawPath(p, fill=int(bool(fill)), stroke=int(bool(stroke)))

def section_label(c, x, y, text, w=None, fontsize=10):
    tw  = c.stringWidth(text, "Helvetica-Bold", fontsize)
    bw  = w or (tw + 28)
    rr(c, x, y, bw, 24, r=12, fill=C_LIGHT)
    c.setFont("Helvetica-Bold", fontsize)
    c.setFillColor(C_DARK)
    c.drawString(x + 14, y + 7, text)
    return bw

def gout_logo(c, x, y, sz=28):
    c.setFont("Helvetica-Bold", sz)
    c.setFillColor(C_DARK)
    c.drawString(x, y, "GOUT.")
    c.setFont("Helvetica", sz * 0.32)
    c.drawString(x + 2, y - sz * 0.30, "GLOBAL OUTSOURCING")


# ─────────────────────────────────────────────────────────────────────────────
# CARGA DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
def load_data():
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    gc     = gspread.authorize(creds)
    sh1    = gc.open_by_key(SHEET1_ID)
    sh2    = gc.open_by_key(SHEET2_ID)

    # ── DIARIO: series diarias ────────────────────────────────────────────────
    diario  = sh1.worksheet("DIARIO").get_all_values()
    header  = diario[5]
    date_cols = [j for j, v in enumerate(header)
                 if v and '/' in v and 'SEM' not in v]
    dates = [header[j] for j in date_cols]

    def daily_ints(row_idx):
        row = diario[row_idx]
        return [parse_int(row[j]) for j in date_cols if row[j]]

    def daily_pct(row_idx):
        row = diario[row_idx]
        return [parse_pct(row[j]) for j in date_cols if row[j]]

    datos_daily  = daily_ints(6)
    ventas_daily = daily_ints(9)
    final_daily  = daily_ints(12)
    efect_serie  = daily_pct(15)

    # ── DATOS total y promedio (hoja DATOS) ───────────────────────────────────
    datos_ws   = sh1.worksheet("DATOS").get_all_values()
    datos_row  = next((r for r in datos_ws if r[1].strip().upper() == 'DATOS'), None)
    datos_total = parse_int(datos_row[34]) if datos_row else 0
    datos_prom  = round(datos_total / max(len([v for v in datos_daily if v > 0]), 1))

    # ── VENTAS BRUTAS (hoja VTAS - Suma total) ────────────────────────────────
    vtas_ws = sh1.worksheet("VTAS").get_all_values()
    suma_row = next((r for r in vtas_ws if r[1].strip() == 'Suma total'), None)
    ventas_total = parse_int(suma_row[-2]) if suma_row else 0
    ventas_prom  = round(ventas_total / max(len([v for v in ventas_daily if v > 0]), 1))

    # ── FINALIZADAS (hoja VTAS - seccion finalizadas) ─────────────────────────
    final_total = 0
    for i, row in enumerate(vtas_ws):
        if row[1].strip().upper() == 'TOTAL' and i > 75:
            val = parse_int(row[-2])
            if val > 0:
                final_total = val
                break
    final_prom = round(final_total / max(len([v for v in final_daily if v > 0]), 1))

    # Efectividad = finalizadas / datos
    efect_total = round((final_total / datos_total * 100), 2) if datos_total else 0.0

    # ── VTAS por asesor ───────────────────────────────────────────────────────
    asesores_vtas = {}
    turno_map     = {}
    for row in vtas_ws[4:19]:
        if row[0] in ('TM', 'TN') and row[1] and row[1] not in ('TM', 'TN', 'Suma total'):
            name = clean_name(row[1])
            try:
                total = int(row[-2])
            except:
                total = 0
            asesores_vtas[name] = total
            turno_map[name]     = row[0]

    # ── DATOS por asesor ──────────────────────────────────────────────────────
    asesores_datos = {}
    for row in datos_ws[4:16]:
        if row[1] and '_UP' in row[1]:
            name = clean_name(row[1])
            asesores_datos[name] = parse_int(row[34])
            if name not in turno_map:
                tn_names = ['Ots', 'Romero', 'Camacho', 'Gonzalez', 'Garcia']
                turno_map[name] = 'TN' if any(n in name for n in tn_names) else 'TM'

    # ── 3HS: datos por hora ───────────────────────────────────────────────────
    hs_raw = sh1.worksheet("3HS").get("AB3:BJ19")
    hora_totals = {}
    for row in hs_raw[1:16]:
        if row and len(row) > 1 and row[1]:
            try:
                hora  = int(row[1])
                total = sum(parse_int(v) for v in row[2:])
                hora_totals[hora] = total
            except:
                pass

    # ── CONVXSKILL: detalle por asesor ────────────────────────────────────────
    cvx_ws  = sh1.worksheet("CONVXSKILL").get_all_values()
    cvx_tm  = []   # lista de dicts por asesor TM
    cvx_tn  = []
    cvx_sub = {}   # subtotales y total

    def parse_cvx_row(row):
        return {
            'nombre':     clean_name(row[3]),
            'turno':      row[2].strip(),
            'datos':      parse_int(row[4]),
            'prom_datos': parse_int(row[5]),
            'horas':      parse_int(row[6]),
            'dias_lab':   parse_int(row[7]),
            'vts_brutas': parse_int(row[8]),
            'conv_brtas': row[9].replace(',', '.').strip(),
            'vts_activ':  parse_int(row[10]),
            'conv_activ': row[11].replace(',', '.').strip(),
        }

    for i, row in enumerate(cvx_ws):
        if len(row) < 12:
            continue
        turno = row[2].strip()
        label = row[3].strip()

        if turno == 'TM' and row[0].strip() and '_UP' in row[0]:
            cvx_tm.append(parse_cvx_row(row))
        elif turno == 'TN' and row[0].strip() and '_UP' in row[0]:
            cvx_tn.append(parse_cvx_row(row))
        elif 'TURNO MA' in label.upper() or 'TURNO MANANA' in label.upper().replace('Ñ','N'):
            cvx_sub['tm'] = parse_cvx_row(row)
            cvx_sub['tm']['nombre'] = label
        elif 'TURNO NOCHE' in label.upper():
            cvx_sub['tn'] = parse_cvx_row(row)
            cvx_sub['tn']['nombre'] = label
        elif 'TOTAL MENSUAL' in label.upper():
            cvx_sub['total'] = parse_cvx_row(row)
            cvx_sub['total']['nombre'] = label

    # ── ASISTENCIA ────────────────────────────────────────────────────────────
    asist = sh2.worksheet("ASISTENCIA").get_all_values()
    team  = []

    def count_p(row):
        return sum(1 for c in row[5:] if str(c).strip().upper() == 'P')

    # TM
    for row in asist[5:10]:
        if row[3] and row[1]:
            team.append({'grupo': 'TM', 'rol': row[1],
                         'nombre': clean_name(row[3]),
                         'horas': row[-1], 'dias_p': count_p(row)})
    # TN
    for row in asist[14:20]:
        if row[3] and row[1]:
            team.append({'grupo': 'TN', 'rol': row[1],
                         'nombre': clean_name(row[3]),
                         'horas': row[-1], 'dias_p': count_p(row)})
    # Supervisores
    for row in asist[26:29]:
        if row[3] and row[1]:
            team.append({'grupo': 'SUP', 'rol': row[1],
                         'nombre': clean_name(row[3]),
                         'horas': row[-1], 'dias_p': count_p(row)})
    # Backoffice
    for row in asist[32:35]:
        if row[3] and row[1]:
            team.append({'grupo': 'BO', 'rol': row[1],
                         'nombre': clean_name(row[3]),
                         'horas': row[-1], 'dias_p': count_p(row)})

    return {
        'dates': dates,
        'datos_daily':  datos_daily,
        'ventas_daily': ventas_daily,
        'final_daily':  final_daily,
        'efect_serie':  efect_serie,
        'datos_total':  datos_total,
        'datos_prom':   datos_prom,
        'ventas_total': ventas_total,
        'ventas_prom':  ventas_prom,
        'final_total':  final_total,
        'final_prom':   final_prom,
        'efect_total':  efect_total,
        'asesores_vtas':  asesores_vtas,
        'asesores_datos': asesores_datos,
        'turno_map':    turno_map,
        'hora_totals':  hora_totals,
        'cvx_tm':       cvx_tm,
        'cvx_tn':       cvx_tn,
        'cvx_sub':      cvx_sub,
        'team':         team,
    }


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 1 — PORTADA
# ─────────────────────────────────────────────────────────────────────────────
def page_cover(c):
    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_PALE);  c.rect(0, 0, 8, H, fill=1, stroke=0)

    c.setFillColor(HexColor('#D0EDD4')); c.circle(W*0.78, H*0.52, 195, fill=1, stroke=0)
    c.setFillColor(C_DARK);             c.circle(W*0.90, H*0.30, 125, fill=1, stroke=0)
    c.setFillColor(HexColor('#A8D5AD')); c.circle(W*0.64, H*0.43,  92, fill=1, stroke=0)

    ty = H - 58
    rr(c, 18, ty, 68, 22, r=11, stroke=C_DARK, lw=1.5)
    c.setFont("Helvetica", 11); c.setFillColor(C_DARK)
    c.drawCentredString(52, ty + 7, MES)
    rr(c, 96, ty, 58, 22, r=11, stroke=C_DARK, lw=1.5)
    c.drawCentredString(125, ty + 7, ANO)

    gout_logo(c, W - 125, H - 48, sz=30)

    c.setFillColor(C_DARK)
    c.setFont("Helvetica", 62);      c.drawString(22, H*0.62, "Reporte")
    c.setFont("Helvetica", 62);      c.drawString(22, H*0.55, "mensual")
    c.setFont("Helvetica-Bold", 90); c.drawString(22, H*0.43, ANO)
    c.setFont("Helvetica", 16);      c.drawString(22, H*0.37, f"Por: {AUTOR_NOMBRE}")


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 2 — CONTENIDO (interactivo)
# ─────────────────────────────────────────────────────────────────────────────
def page_contents(c):
    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_DARK);  c.rect(W-110, H-110, 110, 110, fill=1, stroke=0)

    c.setFont("Helvetica", 52); c.setFillColor(C_DARK)
    c.drawString(40, H-195, "CONTENIDO")

    items = [
        ("Resumen de Resultados",      ANCHOR_RESUMEN),
        ("Reportes Adicionales",       ANCHOR_ADICIONAL),
        ("Detalle por Asesor",         ANCHOR_DETALLE),
        ("Asistencia del Equipo",      ANCHOR_ASISTENCIA),
        ("Recomendaciones",            ANCHOR_RECOMEND),
    ]
    y0, gap = H - 305, 66
    for i, (label, anchor) in enumerate(items):
        y = y0 - i * gap

        # Decoracion
        c.setStrokeColor(C_DARK); c.setLineWidth(1)
        c.line(38, y+16, 78, y+16)
        c.setFillColor(C_DARK); c.circle(78, y+16, 14, fill=1, stroke=0)

        # Etiqueta clickeable
        box_x, box_w, box_h = 96, 248, 28
        rr(c, box_x, y+2, box_w, box_h, r=14, fill=C_LIGHT)
        c.setFont("Helvetica", 13); c.setFillColor(C_DARK)
        c.drawString(box_x + 14, y + 10, label)
        # Flecha indicadora de link
        c.setFont("Helvetica-Bold", 11); c.setFillColor(C_SEC)
        c.drawString(box_x + box_w - 22, y + 10, ">")

        # Numero de pagina como indicador visual
        pg_map = {ANCHOR_RESUMEN: "p.3", ANCHOR_ADICIONAL: "p.4",
                  ANCHOR_DETALLE: "p.5", ANCHOR_ASISTENCIA: "p.6",
                  ANCHOR_RECOMEND: "p.7"}
        c.setFont("Helvetica", 9); c.setFillColor(C_GRAY)
        c.drawString(box_x + box_w + 8, y + 10, pg_map.get(anchor, ""))

        # Link interno PDF clickeable
        c.linkAbsolute("", anchor, Rect=(box_x, y+2, box_x+box_w, y+2+box_h))


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 3 — RESUMEN DE RESULTADOS
# ─────────────────────────────────────────────────────────────────────────────
def chart_daily(datos_daily, ventas_daily, dates):
    n = min(len(datos_daily), len(ventas_daily), len(dates))
    x = np.arange(n)
    labels = [d.split('/')[0] for d in dates[:n]]

    fig, ax = plt.subplots(figsize=(9.5, 4.0))
    fig.patch.set_facecolor('#EAF5EA'); ax.set_facecolor('#EAF5EA')

    w = 0.38
    b1 = ax.bar(x - w/2, datos_daily[:n],  w, label='Datos',         color=M_TEAL, alpha=0.9)
    b2 = ax.bar(x + w/2, ventas_daily[:n], w, label='Ventas Brutas', color=M_DARK, alpha=0.9)

    # Etiquetas en las barras mas altas (cada 5 dias para no saturar)
    for j, rect in enumerate(b1):
        if j % 5 == 0:
            ax.text(rect.get_x() + rect.get_width()/2, rect.get_height() + 2,
                    str(int(rect.get_height())), ha='center', va='bottom',
                    fontsize=6, color=M_MED)

    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=7, rotation=45)
    ax.tick_params(axis='y', labelsize=8)
    ax.legend(fontsize=9, framealpha=0, loc='upper right')
    ax.spines[['top', 'right']].set_visible(False)
    ax.set_ylabel('Cantidad', fontsize=9, color=M_DARK)
    ax.tick_params(colors=M_DARK)
    ax.grid(axis='y', alpha=0.25, color='gray')
    plt.tight_layout(pad=0.5)
    return fig

def chart_turno_pie(asesores_vtas, turno_map):
    tm = sum(v for k, v in asesores_vtas.items() if turno_map.get(k) == 'TM')
    tn = sum(v for k, v in asesores_vtas.items() if turno_map.get(k) == 'TN')

    fig, ax = plt.subplots(figsize=(4, 3.5))
    fig.patch.set_facecolor('#EAF5EA'); ax.set_facecolor('#EAF5EA')

    wedges, texts, autotexts = ax.pie(
        [tm, tn],
        labels=[f'TM\n{tm}', f'TN\n{tn}'],
        colors=[M_TEAL, M_DARK],
        autopct='%1.0f%%', startangle=90,
        textprops={'fontsize': 11},
        wedgeprops={'linewidth': 2, 'edgecolor': 'white'}
    )
    for t in texts:    t.set_color(M_DARK); t.set_fontsize(11)
    for at in autotexts: at.set_fontsize(10); at.set_color('white')
    ax.set_title('Ventas por Turno', fontsize=11, color=M_DARK, pad=6)
    plt.tight_layout(pad=0.5)
    return fig

def chart_efect(efect_serie, dates):
    n = min(len(efect_serie), len(dates))
    x = np.arange(n)
    prom = sum(efect_serie[:n]) / n if n else 0

    fig, ax = plt.subplots(figsize=(5, 2.8))
    fig.patch.set_facecolor('#EAF5EA'); ax.set_facecolor('#EAF5EA')

    ax.plot(x, efect_serie[:n], color=M_DARK, lw=1.8, marker='o', markersize=3)
    ax.fill_between(x, efect_serie[:n], alpha=0.15, color=M_TEAL)
    ax.axhline(y=prom, color=M_TEAL, ls='--', lw=1.2, label=f'Prom. {prom:.1f}%')

    # Etiqueta maximo y minimo
    max_i = int(np.argmax(efect_serie[:n]))
    min_i = int(np.argmin(efect_serie[:n]))
    for idx, prefix in [(max_i, ''), (min_i, '')]:
        ax.annotate(f"{efect_serie[idx]:.1f}%",
                    xy=(idx, efect_serie[idx]),
                    xytext=(0, 6 if idx == max_i else -12),
                    textcoords='offset points',
                    ha='center', fontsize=7, color=M_DARK)

    lbl = [d.split('/')[0] for d in dates[:n]]
    ax.set_xticks(x[::4]); ax.set_xticklabels(lbl[::4], fontsize=7, rotation=45)
    ax.tick_params(axis='y', labelsize=8)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f'{v:.0f}%'))
    ax.legend(fontsize=8, framealpha=0)
    ax.spines[['top', 'right']].set_visible(False)
    ax.set_title('Efectividad diaria', fontsize=10, color=M_DARK)
    ax.tick_params(colors=M_DARK)
    ax.grid(axis='y', alpha=0.2)
    plt.tight_layout(pad=0.5)
    return fig

def page_results(c, d):
    c.bookmarkPage(ANCHOR_RESUMEN)
    c.addOutlineEntry("Resumen de Resultados", ANCHOR_RESUMEN, level=0)

    # Header
    c.setFillColor(C_DARK); c.rect(0, H-95, W, 95, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 19); c.setFillColor(C_WHITE)
    c.drawCentredString(W/2, H-52, "RESUMEN DE RESULTADOS")
    c.setFillColor(C_LIGHT); c.rect(0, 0, 6, H-95, fill=1, stroke=0)

    # KPIs
    kpis = [
        ("DATOS",         str(d['datos_total']),    f"Prom. {d['datos_prom']}/dia"),
        ("VENTAS BRUTAS", str(d['ventas_total']),   f"Prom. {d['ventas_prom']}/dia"),
        ("FINALIZADAS",   str(d['final_total']),    f"Prom. {d['final_prom']}/dia"),
        ("EFECTIVIDAD",   f"{d['efect_total']:.2f}%", "conversion total"),
    ]
    c.setFont("Helvetica-Bold", 11); c.setFillColor(C_DARK)
    c.drawString(18, H-116, "KPI'S CLAVES")

    bw, bh = 126, 78
    gap = (W - 36 - 4*bw) / 3
    by  = H - 205

    for i, (title, val, sub) in enumerate(kpis):
        bx = 18 + i*(bw+gap)
        rr(c, bx, by, bw, bh, r=10, fill=C_WHITE, stroke=C_LIGHT, lw=1.5)
        c.setFont("Helvetica-Bold", 9);  c.setFillColor(C_DARK)
        c.drawCentredString(bx+bw/2, by+bh-17, title)
        c.setFont("Helvetica-Bold", 27); c.drawCentredString(bx+bw/2, by+bh-50, val)
        c.setFont("Helvetica", 8);       c.setFillColor(C_GRAY)
        c.drawCentredString(bx+bw/2, by+bh-64, sub)

    # ── Texto narrativo ───────────────────────────────────────────────────────
    tm_d = sum(v for k,v in d['asesores_datos'].items() if d['turno_map'].get(k)=='TM')
    tn_d = sum(v for k,v in d['asesores_datos'].items() if d['turno_map'].get(k)=='TN')
    tm_v = sum(v for k,v in d['asesores_vtas'].items()  if d['turno_map'].get(k)=='TM')
    tn_v = sum(v for k,v in d['asesores_vtas'].items()  if d['turno_map'].get(k)=='TN')
    mejor_t = "TM (manana)" if tm_v >= tn_v else "TN (noche)"
    narrativa = (
        f"Durante {MES} {ANO} el equipo gestiono {d['datos_total']:,} datos de WhatsApp, "
        f"generando {d['ventas_total']:,} ventas brutas y {d['final_total']:,} finalizadas. "
        f"La efectividad global fue de {d['efect_total']:.2f}%, con el turno {mejor_t} "
        f"liderando en volumen de ventas ({max(tm_v,tn_v)} ventas)."
    )
    words_n = narrativa.split(); lines_n = []; cur_n = ""
    for ww in words_n:
        if len(cur_n) + len(ww) + 1 <= 95:
            cur_n += (" " if cur_n else "") + ww
        else:
            lines_n.append(cur_n); cur_n = ww
    if cur_n: lines_n.append(cur_n)

    rr(c, 14, H-260, W-28, 48, r=6, fill=C_PALE, stroke=C_LIGHT, lw=1)
    c.setFont("Helvetica", 9); c.setFillColor(C_DARK)
    nar_y = H-222
    for ln in lines_n[:3]:
        c.drawString(22, nar_y, ln); nar_y -= 13

    # Grafico diario
    c.setFont("Helvetica-Bold", 10); c.setFillColor(C_DARK)
    c.drawString(18, H-278, "EVOLUCION DIARIA - DATOS Y VENTAS")
    rr(c, 14, H-480, W-28, 194, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)
    img1 = fig_to_img(chart_daily(d['datos_daily'], d['ventas_daily'], d['dates']))
    c.drawImage(img1, 18, H-476, width=W-36, height=186, preserveAspectRatio=True, mask='auto')

    # Turno + Efectividad
    c.setFont("Helvetica-Bold", 10); c.setFillColor(C_DARK)
    c.drawString(18,     H-500, "VENTAS POR TURNO")
    c.drawString(W/2+10, H-500, "EFECTIVIDAD DIARIA")

    half = W/2 - 22
    rr(c, 14,    H-692, half, 184, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)
    rr(c, W/2+6, H-692, half, 184, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)

    img2 = fig_to_img(chart_turno_pie(d['asesores_vtas'], d['turno_map']))
    c.drawImage(img2, 18, H-688, width=half-8, height=176, preserveAspectRatio=True, mask='auto')

    img3 = fig_to_img(chart_efect(d['efect_serie'], d['dates']))
    c.drawImage(img3, W/2+10, H-688, width=half-8, height=176, preserveAspectRatio=True, mask='auto')

    # Tabla TM/TN
    hdrs = ["Turno", "Datos", "Ventas"]
    rows = [("TM - Manana", str(tm_d), str(tm_v)),
            ("TN - Noche",  str(tn_d), str(tn_v))]
    cws  = [100, 55, 55]; rh = 22; tx = 18; ty = H - 718

    for ci, (h, cw) in enumerate(zip(hdrs, cws)):
        cx = tx + sum(cws[:ci])
        rr(c, cx, ty, cw-2, rh-2, r=4, fill=C_DARK)
        c.setFont("Helvetica-Bold", 8); c.setFillColor(C_WHITE)
        c.drawCentredString(cx+cw/2-1, ty+6, h)
    for ri, row in enumerate(rows):
        for ci, (val, cw) in enumerate(zip(row, cws)):
            cx = tx + sum(cws[:ci])
            rr(c, cx, ty-(ri+1)*rh, cw-2, rh-2, r=4,
               fill=C_PALE if ri%2==0 else C_WHITE, stroke=C_LIGHT, lw=0.5)
            c.setFont("Helvetica", 8); c.setFillColor(C_DARK)
            c.drawCentredString(cx+cw/2-1, ty-(ri+1)*rh+6, val)


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 4 — REPORTES ADICIONALES
# ─────────────────────────────────────────────────────────────────────────────
def chart_h_bars(names, values, turno_map, xlabel):
    colors = [M_TEAL if turno_map.get(n)=='TM' else M_DARK for n in names]

    fig, ax = plt.subplots(figsize=(8.5, max(4.5, 0.52 * len(names) + 1.8)))
    fig.patch.set_facecolor('#EAF5EA'); ax.set_facecolor('#EAF5EA')

    bars = ax.barh(names, values, color=colors, height=0.55)
    # Etiquetas con valor
    for bar, val in zip(bars, values):
        ax.text(bar.get_width() + max(values)*0.01, bar.get_y() + bar.get_height()/2,
                str(val), va='center', ha='left', fontsize=9, color=M_DARK,
                fontweight='bold')

    ptm = mpatches.Patch(color=M_TEAL, label='TM Manana')
    ptn = mpatches.Patch(color=M_DARK, label='TN Noche')
    ax.legend(handles=[ptm, ptn], fontsize=8, framealpha=0, loc='lower right')
    ax.spines[['top','right','bottom']].set_visible(False)
    ax.set_xlabel(xlabel, fontsize=9, color=M_DARK)
    ax.tick_params(axis='x', labelsize=8, colors=M_DARK)
    ax.tick_params(axis='y', labelsize=9.5, colors=M_DARK)
    ax.set_xlim(0, max(values) * 1.18)
    ax.grid(axis='x', alpha=0.2)
    fig.subplots_adjust(left=0.30, right=0.95, top=0.95, bottom=0.12)
    return fig

def chart_horas(hora_totals):
    horas  = sorted(hora_totals.keys())
    values = [hora_totals[h] for h in horas]
    max_v  = max(values) if values else 1
    colors = [M_DARK if v == max_v else M_TEAL for v in values]

    fig, ax = plt.subplots(figsize=(9.5, 2.8))
    fig.patch.set_facecolor('#EAF5EA'); ax.set_facecolor('#EAF5EA')

    bars = ax.bar([f'{h}hs' for h in horas], values, color=colors, width=0.68)
    # Etiqueta en cada barra
    for bar, val in zip(bars, values):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + max_v*0.01,
                str(val), ha='center', va='bottom', fontsize=8,
                color=M_DARK, fontweight='bold')

    ax.spines[['top','right']].set_visible(False)
    ax.set_ylabel('Total datos (acum.)', fontsize=9, color=M_DARK)
    ax.tick_params(axis='x', labelsize=8, colors=M_DARK)
    ax.tick_params(axis='y', labelsize=8, colors=M_DARK)
    ax.set_ylim(0, max_v * 1.18)
    ax.grid(axis='y', alpha=0.2)
    plt.tight_layout(pad=0.5)
    return fig

def draw_h_bars_native(c, bx, by, bw, bh, names, values, turno_map, xlabel):
    """Barras horizontales 100% vectoriales — sin pixelacion."""
    n = len(names)
    if n == 0:
        return
    max_val = max(values) if values else 1

    rr(c, bx, by, bw, bh, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)

    ml, mr, mt, mb = 115, 46, 28, 24
    ax_x = bx + ml
    ax_y = by + mb
    ax_w = bw - ml - mr
    ax_h = bh - mt - mb

    slot_h = ax_h / n
    bar_h  = slot_h * 0.60

    # Grid vertical
    c.setStrokeColor(HexColor('#C8DCC8')); c.setLineWidth(0.4)
    for gi in range(1, 5):
        gx = ax_x + ax_w * gi / 4
        c.line(gx, ax_y, gx, ax_y + ax_h)

    # Barras, nombres y valores
    for i, (name, val) in enumerate(zip(names, values)):
        slot_y = ax_y + (n - 1 - i) * slot_h
        bar_y  = slot_y + (slot_h - bar_h) / 2
        bar_w  = max((val / max_val) * ax_w, 3)

        fc = C_TEAL if turno_map.get(name) == 'TM' else C_DARK
        rr(c, ax_x, bar_y, bar_w, bar_h, r=3, fill=fc)

        # Nombre a la izquierda
        c.setFont("Helvetica", 8.5); c.setFillColor(C_DARK)
        c.drawRightString(ax_x - 6, bar_y + bar_h / 2 - 4, name)

        # Valor al final de la barra
        c.setFont("Helvetica-Bold", 9); c.setFillColor(C_DARK)
        c.drawString(ax_x + bar_w + 5, bar_y + bar_h / 2 - 4, str(val))

    # Ticks eje X
    c.setFont("Helvetica", 7.5); c.setFillColor(C_GRAY)
    for gi in range(5):
        gx = ax_x + ax_w * gi / 4
        c.drawCentredString(gx, ax_y - 14, str(int(max_val * gi / 4)))

    # Label eje X
    c.setFont("Helvetica", 8); c.setFillColor(C_GRAY)
    c.drawCentredString(ax_x + ax_w / 2, by + 6, xlabel)

    # Leyenda (arriba derecha)
    leg_x = bx + bw - 108
    leg_y = by + bh - 20
    rr(c, leg_x, leg_y, 10, 8, r=2, fill=C_TEAL)
    c.setFont("Helvetica", 8); c.setFillColor(C_DARK)
    c.drawString(leg_x + 13, leg_y + 1, "TM Manana")
    rr(c, leg_x + 74, leg_y, 10, 8, r=2, fill=C_DARK)
    c.drawString(leg_x + 87, leg_y + 1, "TN Noche")


def page_additional(c, d):
    c.bookmarkPage(ANCHOR_ADICIONAL)
    c.addOutlineEntry("Reportes Adicionales", ANCHOR_ADICIONAL, level=0)

    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_LIGHT); c.rect(0, 0, 6, H, fill=1, stroke=0)
    c.setFillColor(C_DARK);  c.rect(W-85, H-85, 85, 85, fill=1, stroke=0)

    c.setFont("Helvetica", 36); c.setFillColor(C_DARK)
    c.drawString(18, H-68,  "REPORTES")
    c.drawString(18, H-108, "ADICIONALES")

    names_v = list(d['asesores_vtas'].keys())
    vals_v  = list(d['asesores_vtas'].values())
    names_d = list(d['asesores_datos'].keys())
    vals_d  = list(d['asesores_datos'].values())

    # Altura dinamica segun cantidad de asesores (18pt por barra + margenes)
    h_bar_v = max(150, len(names_v) * 18 + 52)
    h_bar_d = max(150, len(names_d) * 18 + 52)

    y1 = H - 138
    section_label(c, 28, y1, "VENTAS BRUTAS POR ASESOR", w=210)
    draw_h_bars_native(c, 14, y1-h_bar_v-10, W-28, h_bar_v,
                       names_v, vals_v, d['turno_map'], 'Ventas Brutas')

    y2 = y1 - (h_bar_v + 36)
    section_label(c, 28, y2, "DATOS (WhatsApp) POR ASESOR", w=220)
    draw_h_bars_native(c, 14, y2-h_bar_d-10, W-28, h_bar_d,
                       names_d, vals_d, d['turno_map'], 'Total Datos')

    y3 = y2 - (h_bar_d + 36)
    section_label(c, 28, y3, "DATOS POR FRANJA HORARIA", w=200)
    rr(c, 14, y3-148, W-28, 138, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)
    img3 = fig_to_img(chart_horas(d['hora_totals']))
    c.drawImage(img3, 18, y3-144, width=W-36, height=130, preserveAspectRatio=True, mask='auto')


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 5 — DETALLE POR ASESOR (formato CONVXSKILL)
# ─────────────────────────────────────────────────────────────────────────────
def draw_cvx_row(c, x, y, row_data, col_widths, rh, bg, bold=False, text_color=None):
    tc = text_color or C_DARK
    for ci, (val, cw) in enumerate(zip(row_data, col_widths)):
        cx = x + sum(col_widths[:ci])
        rr(c, cx, y, cw-1, rh-1, r=2, fill=bg, stroke=C_LIGHT, lw=0.3)
        font = "Helvetica-Bold" if bold else "Helvetica"
        c.setFont(font, 8 if ci > 0 else 9)
        c.setFillColor(tc)
        if ci == 0:
            c.drawString(cx + 4, y + (rh-1)/2 - 4, val)
        else:
            c.drawCentredString(cx + cw/2, y + (rh-1)/2 - 4, val)

def page_detail(c, d):
    c.bookmarkPage(ANCHOR_DETALLE)
    c.addOutlineEntry("Detalle por Asesor", ANCHOR_DETALLE, level=0)

    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_LIGHT); c.rect(0, 0, 6, H, fill=1, stroke=0)
    c.setFillColor(C_DARK);  c.rect(W-85, H-85, 85, 85, fill=1, stroke=0)

    # Titulo
    c.setFont("Helvetica", 36); c.setFillColor(C_DARK)
    c.drawString(18, H-68,  "DETALLE")
    c.drawString(18, H-108, "POR ASESOR")

    # Sub-titulo
    c.setFont("Helvetica", 11); c.setFillColor(C_GRAY)
    c.drawString(18, H-125, f"Acumulado Mensual Upselling Argentina - {MES} {ANO}")

    # ── Tabla ─────────────────────────────────────────────────────────────────
    col_widths = [148, 44, 44, 38, 44, 46, 42, 46, 42]
    headers    = ["ASESOR", "DATOS", "PROM/DIA", "HS", "DIAS LAB.",
                  "VTS BRUTAS", "CONV%", "VTS ACTIV.", "CONV%"]
    total_w    = sum(col_widths)
    tx = (W - total_w) / 2
    ty = H - 155
    rh = 26

    # Fila de encabezados
    for ci, (h, cw) in enumerate(zip(headers, col_widths)):
        cx = tx + sum(col_widths[:ci])
        rr(c, cx, ty, cw-1, rh-1, r=3, fill=C_DARK)
        c.setFont("Helvetica-Bold", 8); c.setFillColor(C_WHITE)
        c.drawCentredString(cx + cw/2, ty + rh/2 - 5, h)
    ty -= rh

    def cvx_vals(asesor):
        return [
            asesor['nombre'],
            str(asesor['datos']),
            str(asesor['prom_datos']),
            str(asesor['horas']),
            str(asesor['dias_lab']),
            str(asesor['vts_brutas']),
            asesor['conv_brtas'],
            str(asesor['vts_activ']),
            asesor['conv_activ'],
        ]

    # ── TM ────────────────────────────────────────────────────────────────────
    # Sub-encabezado TM
    for ci, cw in enumerate(col_widths):
        cx = tx + sum(col_widths[:ci])
        rr(c, cx, ty, cw-1, rh-1, r=3, fill=C_SEC)
    c.setFont("Helvetica-Bold", 9); c.setFillColor(C_WHITE)
    c.drawString(tx + 8, ty + rh/2 - 5, "TURNO MANANA - Natalia Sarcevic")
    ty -= rh

    for ri, asesor in enumerate(d['cvx_tm']):
        bg = C_PALE if ri % 2 == 0 else C_WHITE
        draw_cvx_row(c, tx, ty, cvx_vals(asesor), col_widths, rh, bg)
        ty -= rh

    # Subtotal TM
    if 'tm' in d['cvx_sub']:
        sub = d['cvx_sub']['tm']
        label = f"Subtotal TM ({len(d['cvx_tm'])} asesores)"
        vals  = [label, str(sub['datos']), str(sub['prom_datos']), str(sub['horas']),
                 str(sub['dias_lab']), str(sub['vts_brutas']), sub['conv_brtas'],
                 str(sub['vts_activ']), sub['conv_activ']]
        draw_cvx_row(c, tx, ty, vals, col_widths, rh, C_LIGHT, bold=True)
        ty -= rh

    ty -= 6  # separador

    # ── TN ────────────────────────────────────────────────────────────────────
    for ci, cw in enumerate(col_widths):
        cx = tx + sum(col_widths[:ci])
        rr(c, cx, ty, cw-1, rh-1, r=3, fill=C_SEC)
    c.setFont("Helvetica-Bold", 9); c.setFillColor(C_WHITE)
    c.drawString(tx + 8, ty + rh/2 - 5, "TURNO NOCHE - Giancarlo")
    ty -= rh

    for ri, asesor in enumerate(d['cvx_tn']):
        bg = C_PALE if ri % 2 == 0 else C_WHITE
        draw_cvx_row(c, tx, ty, cvx_vals(asesor), col_widths, rh, bg)
        ty -= rh

    if 'tn' in d['cvx_sub']:
        sub = d['cvx_sub']['tn']
        label = f"Subtotal TN ({len(d['cvx_tn'])} asesores)"
        vals  = [label, str(sub['datos']), str(sub['prom_datos']), str(sub['horas']),
                 str(sub['dias_lab']), str(sub['vts_brutas']), sub['conv_brtas'],
                 str(sub['vts_activ']), sub['conv_activ']]
        draw_cvx_row(c, tx, ty, vals, col_widths, rh, C_LIGHT, bold=True)
        ty -= rh

    ty -= 8  # separador

    # ── Total general ─────────────────────────────────────────────────────────
    if 'total' in d['cvx_sub']:
        tot = d['cvx_sub']['total']
        n_as = len(d['cvx_tm']) + len(d['cvx_tn'])
        label = f"TOTAL MENSUAL ({n_as} asesores)"
        vals  = [label, str(tot['datos']), str(tot['prom_datos']), str(tot['horas']),
                 str(tot['dias_lab']), str(tot['vts_brutas']), tot['conv_brtas'],
                 str(tot['vts_activ']), tot['conv_activ']]
        draw_cvx_row(c, tx, ty, vals, col_widths, rh+2,
                     C_DARK, bold=True, text_color=C_WHITE)
        ty -= (rh + 2)

    # Leyenda columnas
    ty -= 12
    c.setFont("Helvetica", 7.5); c.setFillColor(C_GRAY)
    leyenda = ("VTS BRUTAS = ventas brutas WhatsApp  |  CONV% = conversion  |  "
               "VTS ACTIV. = ventas finalizadas/activadas")
    c.drawCentredString(W/2, ty, leyenda)

    c.setFillColor(C_LIGHT); c.rect(0, 0, W, 12, fill=1, stroke=0)


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 6 — ASISTENCIA DEL EQUIPO
# ─────────────────────────────────────────────────────────────────────────────
def page_attendance(c, d):
    c.bookmarkPage(ANCHOR_ASISTENCIA)
    c.addOutlineEntry("Asistencia del Equipo", ANCHOR_ASISTENCIA, level=0)

    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_DARK);  c.rect(0, H-95, W, 95, fill=1, stroke=0)
    c.setFont("Helvetica-Bold", 20); c.setFillColor(C_WHITE)
    c.drawCentredString(W/2, H-52, "ASISTENCIA DEL EQUIPO")
    c.setFont("Helvetica", 12)
    c.drawCentredString(W/2, H-72, f"{MES} {ANO}")
    c.setFillColor(C_LIGHT); c.rect(0, 0, 6, H-95, fill=1, stroke=0)

    team   = d['team']
    cws    = [28, 55, 185, 55, 60, 60, 70]
    hdrs   = ["#", "ROL", "NOMBRE", "TURNO", "HS/TURNO", "DIAS P", "SUPERVISOR"]
    supmap = {'TM': 'Natalia Sarcevic', 'TN': 'Giancarlo', 'SUP': '-', 'BO': '-'}

    tx = 14; ty = H - 120; rh = 26

    for ci, (h, cw) in enumerate(zip(hdrs, cws)):
        cx = tx + sum(cws[:ci])
        rr(c, cx, ty, cw-1, rh-1, r=3, fill=C_DARK)
        c.setFont("Helvetica-Bold", 8); c.setFillColor(C_WHITE)
        c.drawCentredString(cx+cw/2, ty+9, h)
    ty -= rh

    grupo_prev = None; row_n = 0
    labels_g = {'TM': 'TURNO MANANA', 'TN': 'TURNO NOCHE',
                'SUP': 'SUPERVISORES', 'BO': 'BACKOFFICE'}

    for m in team:
        grupo = m['grupo']
        if grupo != grupo_prev:
            grupo_prev = grupo; row_n = 0
            total_w = sum(cws) - 1
            rr(c, tx, ty, total_w, rh-1, r=3, fill=C_SEC)
            c.setFont("Helvetica-Bold", 9); c.setFillColor(C_WHITE)
            c.drawString(tx+10, ty+9, labels_g.get(grupo, grupo))
            ty -= rh

        row_n += 1
        bg = C_PALE if row_n % 2 == 0 else C_WHITE
        turno_txt = grupo if grupo in ('TM', 'TN') else '-'
        dias_p_txt = str(m['dias_p']) if m['dias_p'] else '-'
        row_vals = [str(row_n), m['rol'], m['nombre'], turno_txt,
                    (str(m['horas'])+'hs') if m['horas'] else '-',
                    dias_p_txt, supmap.get(grupo, '-')]

        for ci, (val, cw) in enumerate(zip(row_vals, cws)):
            cx = tx + sum(cws[:ci])
            rr(c, cx, ty, cw-1, rh-1, r=3, fill=bg, stroke=C_LIGHT, lw=0.4)
            c.setFont("Helvetica", 8); c.setFillColor(C_DARK)
            if ci == 2:
                c.drawString(cx+5, ty+9, val)
            else:
                c.drawCentredString(cx+cw/2, ty+9, val)
        ty -= rh

    # Nota al pie
    ty -= 8
    c.setFont("Helvetica", 8); c.setFillColor(C_GRAY)
    c.drawString(tx, ty, "DIAS P = cantidad de dias marcados como Presente (P) en el mes")
    c.setFillColor(C_LIGHT); c.rect(0, 0, W, 12, fill=1, stroke=0)


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 7 — CIERRE
# ─────────────────────────────────────────────────────────────────────────────
def page_closing(c):
    c.setFillColor(C_WHITE); c.rect(0, H/2, W, H/2, fill=1, stroke=0)
    gout_logo(c, W/2 - 65, H*0.77, sz=34)

    c.setFont("Helvetica-Bold", 36); c.setFillColor(C_DARK)
    c.drawCentredString(W/2, H*0.72, "SEGUIMOS")
    c.setFont("Helvetica", 36)
    c.drawCentredString(W/2, H*0.685, "CRECIENDO")

    cy = H*0.645
    items = [
        ("Arg.",   "Argentina"),
        ("Tel.",   AUTOR_TEL),
        ("in",     AUTOR_LINKEDIN),
        ("git",    AUTOR_GITHUB),
    ]
    for sym, val in items:
        c.setFont("Helvetica-Bold", 10); c.setFillColor(C_DARK)
        c.drawString(W/2 - 130, cy, sym)
        c.setFont("Helvetica", 10)
        c.drawString(W/2 - 100, cy, val)
        cy -= 19

    c.setFillColor(HexColor('#B2DFDB')); c.rect(0, 0, W, H/2, fill=1, stroke=0)
    c.setFillColor(HexColor('#80CBC4')); c.circle(W*0.35, H*0.25, 90, fill=1, stroke=0)
    c.setFillColor(HexColor('#4DB6AC')); c.circle(W*0.35, H*0.25, 60, fill=1, stroke=0)
    c.setFillColor(HexColor('#26A69A')); c.circle(W*0.35, H*0.25, 35, fill=1, stroke=0)


# ─────────────────────────────────────────────────────────────────────────────
# PAGINA 7 — RECOMENDACIONES
# ─────────────────────────────────────────────────────────────────────────────
def build_recommendations(d):
    """Genera recomendaciones automaticas basadas en los datos del mes."""
    pos = []  # lo que funciono bien
    imp = []  # para mejorar

    cvx_all = d['cvx_tm'] + d['cvx_tn']

    # --- Conversion global ---
    efect = d['efect_total']
    if efect >= 16:
        pos.append(f"Efectividad global solida: {efect:.2f}% de conversion de datos a ventas finalizadas.")
    else:
        imp.append(f"Efectividad global del {efect:.2f}% — hay margen para subir la tasa de conversion.")

    # --- Mejor asesor por ventas ---
    if d['asesores_vtas']:
        top_v = max(d['asesores_vtas'], key=d['asesores_vtas'].get)
        top_v_val = d['asesores_vtas'][top_v]
        pos.append(f"Mayor volumen de ventas: {top_v} con {top_v_val} ventas brutas.")

    # --- Mejor conversion (CONV% VTAS BRUTAS) ---
    if cvx_all:
        best_conv = max(cvx_all, key=lambda x: parse_pct(x['conv_brtas']))
        pos.append(
            f"Mejor tasa de conversion: {best_conv['nombre']} con {best_conv['conv_brtas']} "
            f"({best_conv['vts_brutas']} ventas / {best_conv['datos']} datos)."
        )
        worst_conv = min(cvx_all, key=lambda x: parse_pct(x['conv_brtas']))
        if parse_pct(worst_conv['conv_brtas']) < 10:
            imp.append(
                f"Conversion mas baja: {worst_conv['nombre']} con {worst_conv['conv_brtas']}. "
                f"Revisar calidad de datos y script de ventas."
            )

    # --- Asesor con mas datos pero baja conversion ---
    if cvx_all:
        high_data_low_conv = [a for a in cvx_all
                              if a['datos'] > 500 and parse_pct(a['conv_brtas']) < 14]
        for a in high_data_low_conv:
            imp.append(
                f"{a['nombre']} tiene alto volumen de datos ({a['datos']}) "
                f"pero conversion baja ({a['conv_brtas']}). Oportunidad de mejora en cierre."
            )

    # --- Franja horaria pico ---
    if d['hora_totals']:
        hora_pico = max(d['hora_totals'], key=d['hora_totals'].get)
        hora_baja = min(d['hora_totals'], key=d['hora_totals'].get)
        pos.append(
            f"Franja horaria mas productiva: {hora_pico}hs "
            f"({d['hora_totals'][hora_pico]} datos acumulados en el mes)."
        )
        imp.append(
            f"Franja de menor actividad: {hora_baja}hs "
            f"({d['hora_totals'][hora_baja]} datos). Evaluar refuerzo o redistribucion de carga."
        )

    # --- Diferencia TM vs TN ---
    tm_v = sum(v for k, v in d['asesores_vtas'].items() if d['turno_map'].get(k) == 'TM')
    tn_v = sum(v for k, v in d['asesores_vtas'].items() if d['turno_map'].get(k) == 'TN')
    tm_d = sum(v for k, v in d['asesores_datos'].items() if d['turno_map'].get(k) == 'TM')
    tn_d = sum(v for k, v in d['asesores_datos'].items() if d['turno_map'].get(k) == 'TN')

    if tm_d > 0 and tn_d > 0:
        tm_conv = tm_v / tm_d * 100
        tn_conv = tn_v / tn_d * 100
        mejor_turno = "TM" if tm_conv > tn_conv else "TN"
        pos.append(
            f"Turno {mejor_turno} con mayor efectividad de conversion "
            f"(TM: {tm_conv:.1f}% | TN: {tn_conv:.1f}%)."
        )

    # --- Volumen total ---
    pos.append(
        f"Volumen total del mes: {d['datos_total']} datos gestionados con "
        f"{d['ventas_total']} ventas brutas y {d['final_total']} finalizadas."
    )

    return pos[:5], imp[:5]  # maximo 5 de cada tipo


def page_recommendations(c, d):
    c.bookmarkPage(ANCHOR_RECOMEND)
    c.addOutlineEntry("Recomendaciones", ANCHOR_RECOMEND, level=0)

    c.setFillColor(C_WHITE); c.rect(0, 0, W, H, fill=1, stroke=0)
    c.setFillColor(C_DARK);  c.rect(0, H-95, W, 95, fill=1, stroke=0)
    c.setFillColor(C_LIGHT); c.rect(0, 0, 6, H-95, fill=1, stroke=0)

    c.setFont("Helvetica-Bold", 19); c.setFillColor(C_WHITE)
    c.drawCentredString(W/2, H-52, "ANALISIS Y RECOMENDACIONES")
    c.setFont("Helvetica", 11)
    c.drawCentredString(W/2, H-72, f"{MES} {ANO}")

    positivos, mejoras = build_recommendations(d)

    margin = 30
    col_w  = (W - margin*3) / 2
    box_h_unit = 70

    # ── Columna izquierda: Lo que funciono bien ───────────────────────────────
    cx_l = margin
    cy   = H - 115
    section_label(c, cx_l, cy, "LO QUE FUNCIONO BIEN", w=col_w - 10)
    cy  -= 32

    for item in positivos:
        # Calcular altura necesaria (wrap de texto)
        words   = item.split()
        lines   = []
        current = ""
        max_chars = 44
        for w in words:
            if len(current) + len(w) + 1 <= max_chars:
                current += (" " if current else "") + w
            else:
                lines.append(current); current = w
        if current: lines.append(current)
        bh = max(box_h_unit, 22 + len(lines) * 14)

        rr(c, cx_l, cy - bh, col_w, bh, r=8, fill=C_PALE, stroke=C_LIGHT, lw=1)
        # Circulo verde
        c.setFillColor(C_MED)
        c.circle(cx_l + 18, cy - bh/2, 8, fill=1, stroke=0)
        # Signo +
        c.setFont("Helvetica-Bold", 10); c.setFillColor(C_WHITE)
        c.drawCentredString(cx_l + 18, cy - bh/2 - 4, "+")
        # Texto
        c.setFont("Helvetica", 9); c.setFillColor(C_DARK)
        for li, line in enumerate(lines):
            c.drawString(cx_l + 34, cy - 18 - li*14, line)
        cy -= (bh + 8)

    # ── Columna derecha: Para mejorar ─────────────────────────────────────────
    cx_r = margin * 2 + col_w
    cy   = H - 115
    section_label(c, cx_r, cy, "AREAS DE MEJORA", w=col_w - 10)
    cy  -= 32

    for item in mejoras:
        words   = item.split()
        lines   = []
        current = ""
        max_chars = 44
        for w in words:
            if len(current) + len(w) + 1 <= max_chars:
                current += (" " if current else "") + w
            else:
                lines.append(current); current = w
        if current: lines.append(current)
        bh = max(box_h_unit, 22 + len(lines) * 14)

        rr(c, cx_r, cy - bh, col_w, bh, r=8,
           fill=HexColor('#FFF8E1'), stroke=HexColor('#FFE082'), lw=1)
        # Circulo naranja
        c.setFillColor(HexColor('#F57C00'))
        c.circle(cx_r + 18, cy - bh/2, 8, fill=1, stroke=0)
        # Signo !
        c.setFont("Helvetica-Bold", 10); c.setFillColor(C_WHITE)
        c.drawCentredString(cx_r + 18, cy - bh/2 - 4, "!")
        # Texto
        c.setFont("Helvetica", 9); c.setFillColor(C_DARK)
        for li, line in enumerate(lines):
            c.drawString(cx_r + 34, cy - 18 - li*14, line)
        cy -= (bh + 8)

    c.setFillColor(C_LIGHT); c.rect(0, 0, W, 12, fill=1, stroke=0)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    print("Cargando datos de Google Sheets...")
    d = load_data()
    print(f"  Datos: {d['datos_total']} | Ventas: {d['ventas_total']} | "
          f"Finalizadas: {d['final_total']} | Efectividad: {d['efect_total']:.2f}%")
    print(f"  CVX TM: {len(d['cvx_tm'])} asesores | TN: {len(d['cvx_tn'])} asesores")

    print("Generando PDF...")
    cv = pdf_canvas.Canvas(OUTPUT, pagesize=A4)

    page_cover(cv);                cv.showPage()
    page_contents(cv);             cv.showPage()
    page_results(cv, d);           cv.showPage()
    page_additional(cv, d);        cv.showPage()
    page_detail(cv, d);            cv.showPage()
    page_attendance(cv, d);        cv.showPage()
    page_recommendations(cv, d);   cv.showPage()
    page_closing(cv);              cv.showPage()

    cv.save()
    print(f"OK Reporte generado: {OUTPUT}")

if __name__ == "__main__":
    main()
