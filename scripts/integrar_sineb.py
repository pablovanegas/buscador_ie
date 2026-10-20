#!/usr/bin/env python3
"""Integra la base SINEB (MinEducación, «Buscando Colegio») al buscador.

Qué hace:
  1. Cruza las instituciones del HTML (DATA) con la base SINEB por código DANE.
     El HTML usa el código de la sede principal (14 dígitos = DANE institución + "01"),
     así que la llave de cruce son los primeros 12 dígitos.
  2. Inserta en index.html una variable `SINEB` (una línea) con la oferta educativa
     de cada institución que YA está en el HTML: jornadas, grados, modelos flexibles,
     inclusión, sedes y datos de contacto oficiales. No agrega instituciones nuevas.
  3. Exporta a `datos/` dos libros de Excel:
       - instituciones_no_incluidas.xlsx: las instituciones de la base que NO están en el HTML.
       - reporte_cruce_html_sineb.xlsx: calidad del cruce y diferencias de contacto.

Uso:  python3 scripts/integrar_sineb.py [ruta_excel]
Es idempotente: si `var SINEB=` ya existe en index.html, la reemplaza.
"""
import json
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
HTML = ROOT / "index.html"
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "datos/fuente/colegios_medellin_sineb.xlsx"
OUT = ROOT / "datos"

# ── utilidades de texto ──────────────────────────────────────────────
MINUS = {"de", "del", "la", "las", "los", "el", "y", "e", "a", "en", "con", "por", "para", "al", "o"}
SIGLAS = {"UCN", "CLEI", "PFC", "SER", "IE", "CE", "I", "II", "III"}
# Tildes que la fuente omite
ACENTOS = {"Promocion": "Promoción", "Paralisis": "Parálisis", "Sistemica": "Sistémica", "Transtorno": "Trastorno"}


def tc(s):
    """Tipo título en español respetando conectores y siglas."""
    if not isinstance(s, str) or not s.strip():
        return ""
    out = []
    for i, w in enumerate(s.strip().lower().split()):
        if w.upper() in SIGLAS:
            out.append(w.upper())
        elif i and w in MINUS:
            out.append(w)
        else:
            w = w[:1].upper() + w[1:]
            out.append(ACENTOS.get(w, w))
    return " ".join(out)


def norm(s):
    s = unicodedata.normalize("NFD", str(s or "").lower())
    return re.sub(r"[^a-z0-9]", "", s.encode("ascii", "ignore").decode())


def split_list(v, sep=","):
    if not isinstance(v, str):
        return []
    return [p.strip() for p in v.split(sep) if p.strip()]


def phones(v):
    """Separa teléfonos escritos de mil formas: '305 444 03 61', '2552996 2853667', '(604) 429 13 50 - 300...'."""
    if not isinstance(v, str):
        return []
    res = []
    for part in re.split(r"\s*[-/,;|]\s*", v):
        d = re.sub(r"\D", "", part)
        if len(d) in (7, 10):
            res.append(d)
        elif d:
            res += [t for t in (re.sub(r"\D", "", x) for x in part.split()) if len(t) in (7, 10)]
    seen, out = set(), []
    for p in res:
        if p[-7:] not in seen:
            seen.add(p[-7:])
            out.append(p)
    return out


def emails(v):
    return [e.lower() for e in re.findall(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", str(v or ""))]


def grades(v):
    return sorted({int(g) for g in split_list(str(v)) if re.fullmatch(r"-?\d+", g)})


# Orden fijo de jornadas (igual en filtros, tarjetas y estadísticas)
# La jornada NOCTURNA se excluye a propósito: el buscador no publica esa información.
JORNADAS = [("MAÑANA", "Mañana"), ("TARDE", "Tarde"), ("ÚNICA", "Única"),
            ("COMPLETA", "Completa"), ("FIN DE SEMANA", "Fin de semana")]
JOR_MAP = dict(JORNADAS)


def jornadas(v):
    got = {p.upper() for p in split_list(v)}
    return [lbl for k, lbl in JORNADAS if k in got]


def modelos_flexibles(v):
    return [tc(m) for m in split_list(v) if m.upper() != "EDUCACIÓN TRADICIONAL"]


# ── carga ────────────────────────────────────────────────────────────
def load_html_data():
    html = HTML.read_text(encoding="utf-8")
    line = next(l for l in html.split("\n") if l.startswith("var DATA="))
    return html, json.loads(line[len("var DATA="):].rstrip(";"))


def main():
    xl = pd.ExcelFile(SRC)
    cons = xl.parse("Consolidado")
    filt = xl.parse("Filtros")
    sedes = xl.parse("Sedes")
    for df, col in ((cons, "Código DANE"), (filt, "Código DANE"), (sedes, "Código DANE institución")):
        df["_k"] = df[col].astype(str)

    html, data = load_html_data()
    html_by_key = {r["dane"][:12]: r for r in data}
    assert len(html_by_key) == len(data), "Llaves DANE duplicadas en el HTML"

    cons_k = cons.set_index("_k")
    filt_k = filt.set_index("_k")
    matched = [k for k in html_by_key if k in cons_k.index]
    missing_html = [k for k in html_by_key if k not in cons_k.index]
    print(f"HTML: {len(data)} | SINEB: {len(cons)} | cruzan: {len(matched)} | HTML sin SINEB: {len(missing_html)}")

    # ── registro compacto por institución ───────────────────────────
    sineb = {}
    for k in matched:
        c, f = cons_k.loc[k], filt_k.loc[k]
        s_rows = sedes[sedes["_k"] == k].sort_values("Código sede")
        sed = []
        for _, s in s_rows.iterrows():
            sed.append({
                "c": str(s["Código sede"]),
                "n": tc(s["Nombre sede"]),
                "dir": s["Dirección"] if isinstance(s["Dirección"], str) else "",
                "tel": phones(s["Teléfono"]),
                "z": tc(s["Zona"]),
                "e": tc(s["Estado sede"]),
                "jor": jornadas(s["Jornadas"]),
                "gr": grades(s["Grados (códigos)"]),
                "mf": modelos_flexibles(s["Modelos flexibles"]),
            })
        esp = [tc(e) for e in split_list(c["Especialidad"]) if e not in ("OTRO", "NO APLICA")]
        rec = {
            "tipo": "CE" if c["Tipo"] == "CENTRO EDUCATIVO" else "IE",
            "car": tc(c["Carácter"]).replace("/academico", "/Académico") if c["Carácter"] != "NO APLICA" else "",
            "esp": esp,
            "gen": tc(c["Género"]),
            "zona": [tc(z) for z in split_list(c["Zona"])],
            "prest": "Concesión" if c["Prestador del servicio"] == "CONCESION" else "Oficial",
            "jor": [lbl for key, lbl in JORNADAS if f.get("JOR · " + key) == "Sí"],
            "gr": grades(c["Grados (códigos)"]),
            "mf": modelos_flexibles(c["Modelos flexibles"]),
            "adu": c["Atiende adultos"] == "Sí",
            "disc": [tc(d) for d in split_list(c["Discapacidades"])],
            "cap": [tc(d) for d in split_list(c["Capacidades excepcionales"])],
            "etn": [tc(d) for d in split_list(c["Etnias"])],
            "idi": [tc(d) for d in split_list(c["Idiomas"])],
            "rec": tc(re.sub(r"\.$", "", str(c["Rector"]).strip())) if isinstance(c["Rector"], str) else "",
            "mail": emails(c["Correo electrónico"]),
            "tel": phones(c["Teléfono"]),
            "nsed": int(c["Sedes activas"]),
            "sedes": sed,
            "act": str(c["Actualizado"])[:10],
        }
        # Elimina listas vacías para aligerar el HTML
        sineb[html_by_key[k]["dane"]] = {a: b for a, b in rec.items() if b not in ([], "", None)}

    # ── inserta / reemplaza la línea SINEB en el HTML ───────────────
    leeme = " ".join(str(v) for v in xl.parse("Leeme").iloc[:, 0])
    gen = re.search(r"Generado:\s*(\d{4}-\d{2}-\d{2})", leeme)
    meta = {"fuente": "SINEB – MinEducación (Buscando Colegio)", "fecha": gen.group(1) if gen else ""}
    line = "var SINEB=" + json.dumps(sineb, ensure_ascii=False, separators=(",", ":")) + ";"
    meta_line = "var SINEB_META=" + json.dumps(meta, ensure_ascii=False) + ";"
    lines = html.split("\n")
    lines = [l for l in lines if not l.startswith("var SINEB=") and not l.startswith("var SINEB_META=")]
    i = next(n for n, l in enumerate(lines) if l.startswith("var DATA="))
    lines[i + 1:i + 1] = [line, meta_line]
    HTML.write_text("\n".join(lines), encoding="utf-8")
    print(f"index.html actualizado: SINEB con {len(sineb)} instituciones ({len(line)/1024:.0f} KB)")

    # ── Excel 1: instituciones de la base que NO están en el HTML ───
    keys_html = set(html_by_key)
    no_inc = cons[~cons["_k"].isin(keys_html)].drop(columns="_k")
    resumen = (no_inc.groupby(["Sector", "Tipo"]).size().rename("Instituciones").reset_index()
               .sort_values(["Sector", "Instituciones"], ascending=[True, False]))
    vacias = [c for c in no_inc.columns if no_inc[c].isna().all()]
    sheets = {
        "Resumen": resumen,
        "Instituciones": no_inc.drop(columns=vacias),
        "Filtros": filt[~filt["_k"].isin(keys_html)].drop(columns="_k"),
        "Sedes": sedes[~sedes["_k"].isin(keys_html)].drop(columns="_k"),
    }
    for name in ("Modelos", "Niveles_Grados", "Jornadas", "Tarifas", "Tarifas_por_grado"):
        if name in xl.sheet_names:
            df = xl.parse(name)
            sheets[name] = df[~df["Código DANE"].astype(str).isin(keys_html)]
    notas = pd.DataFrame({"Notas": [
        f"Fuente: {SRC.name} (SINEB – MinEducación, consulta 2026-10-16).",
        f"Instituciones en la base: {len(cons)}. En el buscador (HTML): {len(matched)}. No incluidas: {len(no_inc)}.",
        "Cruce por código DANE de la institución (12 dígitos). El HTML usa el código de la sede principal (DANE + '01').",
        f"Columnas sin ningún dato en estas instituciones (omitidas en 'Instituciones'): {', '.join(vacias) or 'ninguna'}.",
        "Las instituciones NO OFICIALES (privadas) incluyen tarifas cuando el colegio las publica en SINEB.",
        "Estas instituciones no tienen coordenadas en la base: para ubicarlas en el mapa hay que geocodificar su dirección.",
    ]})
    write_xlsx(OUT / "instituciones_no_incluidas.xlsx", {"Leeme": notas, **sheets})
    print(f"datos/instituciones_no_incluidas.xlsx: {len(no_inc)} instituciones")

    # ── Excel 2: reporte del cruce y diferencias de contacto ────────
    rows = []
    for k in matched:
        h, c = html_by_key[k], cons_k.loc[k]
        h_tel = {p[-7:] for p in phones(h.get("tel", ""))}
        s_tel = {p[-7:] for p in phones(c["Teléfono"])}
        h_mail = set(emails(h.get("correo", "")))
        s_mail = set(emails(c["Correo electrónico"]))
        h_rec = {norm(w) for w in (h.get("rector") or "").split()}
        s_rec = {norm(w) for w in str(c["Rector"]).split()}
        rows.append({
            "DANE (HTML)": h["dane"], "DANE institución": k,
            "Nombre (HTML)": h["nombre"], "Nombre (SINEB)": c["Nombre"],
            "Nombre coincide": "Sí" if norm(h["nombre"]) == norm(c["Nombre"]) else "No",
            "Rector (HTML)": h.get("rector", ""), "Rector (SINEB)": c["Rector"],
            "Rector coincide": "Sí" if h_rec and h_rec <= s_rec else "No",
            "Teléfono (HTML)": h.get("tel", ""), "Teléfono (SINEB)": c["Teléfono"],
            "Teléfono coincide": "Sí" if h_tel & s_tel else "No",
            "Correo (HTML)": h.get("correo", ""), "Correo (SINEB)": c["Correo electrónico"],
            "Correo coincide": "Sí" if h_mail & s_mail else "No",
            "Dirección (HTML)": h.get("dir", ""), "Dirección (SINEB)": c["Dirección"],
            "Sedes activas": int(c["Sedes activas"]),
        })
    cruce = pd.DataFrame(rows)
    stats = pd.DataFrame([
        {"Indicador": "Instituciones en el HTML", "Valor": len(data)},
        {"Indicador": "Instituciones en SINEB (Medellín)", "Valor": len(cons)},
        {"Indicador": "Cruzan por código DANE", "Valor": len(matched)},
        {"Indicador": "HTML sin registro en SINEB", "Valor": len(missing_html)},
        {"Indicador": "SINEB no incluidas en el HTML", "Valor": len(no_inc)},
        *[{"Indicador": f"{c.replace(' coincide', '')} distinto (HTML vs SINEB)", "Valor": int((cruce[c] == "No").sum())}
          for c in ("Nombre coincide", "Rector coincide", "Teléfono coincide", "Correo coincide")],
    ])
    write_xlsx(OUT / "reporte_cruce_html_sineb.xlsx", {"Resumen": stats, "Cruce": cruce})
    print("datos/reporte_cruce_html_sineb.xlsx")
    print(stats.to_string(index=False))


def write_xlsx(path, sheets):
    head = PatternFill("solid", fgColor="1A1714")
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
            ws = w.sheets[name[:31]]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = head
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            ws.freeze_panes = "A2"
            if len(df):
                ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, 1):
                width = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(max(width, 10), 60)


if __name__ == "__main__":
    main()
