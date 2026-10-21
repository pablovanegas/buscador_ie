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

Privacidad: no se publica ningún dato personal. Se eliminan los nombres de rectores, los
correos no institucionales (Gmail, Hotmail, Yahoo…) y los números de celular. Solo quedan
teléfonos fijos y correos de dominios .gov.co / .edu.co. Esto aplica a la página y a todos los Excel.

Uso:  python3 scripts/integrar_sineb.py [ruta_excel]
      python3 scripts/integrar_sineb.py --sanear-fuente [ruta_excel]   (reescribe la fuente sin datos personales)
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
ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
SRC = Path(ARGS[0]) if ARGS else ROOT / "datos/fuente/colegios_medellin_sineb.xlsx"
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


# ── privacidad ───────────────────────────────────────────────────────
# Buzones con el nombre de una persona, aunque estén en un dominio institucional
# (las siglas de la institución, como mjb = Manuel J. Betancur, sí son institucionales)
CORREOS_PERSONALES = {"alexandralatorre", "diego.garcia"}


def es_institucional(correo):
    """Solo dominios institucionales (.gov.co de la Alcaldía o .edu.co de la propia institución)
    y buzones de cargo o de la institución, no de una persona."""
    usuario, _, dom = correo.lower().rpartition("@")
    if usuario in CORREOS_PERSONALES:
        return False
    return dom.endswith(".gov.co") or dom.endswith(".edu.co")


def es_fijo(tel):
    """Teléfono fijo: 7 dígitos o 10 con indicativo 60x. Los celulares (3xx…) se excluyen."""
    return len(tel) == 7 or (len(tel) == 10 and tel.startswith("60"))


def correos_inst(v):
    return [e for e in emails(v) if es_institucional(e)]


def fijos(v):
    return [t for t in phones(v) if es_fijo(t)]


def sanear_df(df):
    """Quita rectores, correos personales y celulares de cualquier hoja de la base."""
    df = df.drop(columns=[c for c in df.columns if "rector" in c.lower()])
    for c in df.columns:
        if c.lower().startswith("correo"):
            df[c] = df[c].map(lambda v: " ".join(correos_inst(v)) or None)
        elif c.lower().startswith("teléfono"):
            df[c] = df[c].map(lambda v: " - ".join(fijos(v)) or None)
    return df


def sanear_data(data):
    """Datos originales del buscador: sin rector y solo teléfonos fijos y correos institucionales."""
    for r in data:
        r.pop("rector", None)
        r["tel"] = " - ".join(fijos(r.get("tel", "")))
        r["correo"] = " ".join(correos_inst(r.get("correo", "")))
    return data


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
    # Población de cada modelo flexible (fila a nivel de institución): ADULTOS o NIÑOS Y JOVENES
    mods = xl.parse("Modelos")
    mods = mods[mods["Código sede"].astype(str).str.startswith("(") & (mods["Tipo de modelo"] == "Flexible")]
    mod_adultos = {str(k): sorted({tc(m) for m in g["Modelo"]}) for k, g in mods[mods["Población"] == "ADULTOS"].groupby("Código DANE")}
    for df, col in ((cons, "Código DANE"), (filt, "Código DANE"), (sedes, "Código DANE institución")):
        df["_k"] = df[col].astype(str)

    html, data = load_html_data()
    data = sanear_data(data)
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
                "tel": fijos(s["Teléfono"]),
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
            "mfad": mod_adultos.get(k, []),  # modelos flexibles dirigidos a adultos
            "disc": [tc(d) for d in split_list(c["Discapacidades"])],
            "cap": [tc(d) for d in split_list(c["Capacidades excepcionales"])],
            "etn": [tc(d) for d in split_list(c["Etnias"])],
            "idi": [tc(d) for d in split_list(c["Idiomas"])],
            "mail": correos_inst(c["Correo electrónico"]),
            "tel": fijos(c["Teléfono"]),
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
    lines[i] = "var DATA=" + json.dumps(data, ensure_ascii=False) + ";"
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
        "Instituciones": sanear_df(no_inc.drop(columns=vacias)),
        "Filtros": sanear_df(filt[~filt["_k"].isin(keys_html)].drop(columns="_k")),
        "Sedes": sanear_df(sedes[~sedes["_k"].isin(keys_html)].drop(columns="_k")),
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
        "Privacidad: se omiten los nombres de rectores, los correos no institucionales y los números de celular.",
    ]})
    write_xlsx(OUT / "instituciones_no_incluidas.xlsx", {"Leeme": notas, **sheets})
    print(f"datos/instituciones_no_incluidas.xlsx: {len(no_inc)} instituciones")

    # ── Excel 2: reporte del cruce y diferencias de contacto ────────
    rows = []
    for k in matched:
        h, c = html_by_key[k], cons_k.loc[k]
        h_tel = {p[-7:] for p in fijos(h.get("tel", ""))}
        s_tel = {p[-7:] for p in fijos(c["Teléfono"])}
        h_mail = set(correos_inst(h.get("correo", "")))
        s_mail = set(correos_inst(c["Correo electrónico"]))
        rows.append({
            "DANE (HTML)": h["dane"], "DANE institución": k,
            "Nombre (HTML)": h["nombre"], "Nombre (SINEB)": c["Nombre"],
            "Nombre coincide": "Sí" if norm(h["nombre"]) == norm(c["Nombre"]) else "No",
            "Teléfono (HTML)": h.get("tel", ""), "Teléfono (SINEB)": " - ".join(sorted(fijos(c["Teléfono"]))),
            "Teléfono coincide": "Sí" if h_tel & s_tel else "No",
            "Correo (HTML)": h.get("correo", ""), "Correo (SINEB)": " ".join(correos_inst(c["Correo electrónico"])),
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
          for c in ("Nombre coincide", "Teléfono coincide", "Correo coincide")],
    ])
    write_xlsx(OUT / "reporte_cruce_html_sineb.xlsx", {"Resumen": stats, "Cruce": cruce})
    print("datos/reporte_cruce_html_sineb.xlsx")
    print(stats.to_string(index=False))

    exportar_clei(cons, sedes, html_by_key, mod_adultos)


# ── Excel 3: instituciones con CLEI activo ───────────────────────────
def cleis(gr):
    return [g for g in gr if 20 <= g <= 26]


def clei_txt(gr):
    """'CLEI 1 a 6' si son seguidos; 'CLEI 1, 3 y 5' si hay saltos."""
    n = [g - 20 for g in cleis(gr)]
    if not n:
        return ""
    if len(n) == 1:
        return f"CLEI {n[0]}"
    if n == list(range(n[0], n[-1] + 1)):
        return f"CLEI {n[0]} a {n[-1]}"
    return "CLEI " + ", ".join(map(str, n[:-1])) + f" y {n[-1]}"


def exportar_clei(cons, sedes, html_by_key, mod_adultos):
    """Una fila por institución con CLEI, con columnas que distinguen fin de semana e inclusión.
    La jornada nocturna se omite (decisión del proyecto) y los contactos son solo institucionales."""
    si_no = lambda b: "Sí" if b else "No"
    filas, filas_sede = [], []
    for _, c in cons.iterrows():
        gr = grades(c["Grados (códigos)"])
        if not cleis(gr):
            continue
        k = c["_k"]
        h = html_by_key.get(k)
        sd = sedes[sedes["_k"] == k]
        jor = jornadas(c["Jornadas (ficha)"])
        # sedes donde se ofrece CLEI y fin de semana a la vez
        sedes_fs = []
        for _, x in sd.iterrows():
            xgr, xjor = grades(x["Grados (códigos)"]), jornadas(x["Jornadas"])
            filas_sede.append({
                "DANE institución": k, "Institución": tc(c["Nombre"]), "En el buscador": si_no(h),
                "Código sede": str(x["Código sede"]), "Sede": tc(x["Nombre sede"]),
                "Dirección": x["Dirección"] if isinstance(x["Dirección"], str) else "",
                "Teléfono fijo": " - ".join(fijos(x["Teléfono"])), "Zona": tc(x["Zona"]),
                "Jornadas": ", ".join(xjor), "Ciclos CLEI": clei_txt(xgr),
                "Fin de semana": si_no("Fin de semana" in xjor),
                "CLEI y fin de semana en la sede": si_no(cleis(xgr) and "Fin de semana" in xjor),
            })
            if cleis(xgr) and "Fin de semana" in xjor:
                sedes_fs.append(tc(x["Nombre sede"]))
        rural = "RURAL" in str(c["Zona"]).upper() or sd["Zona"].astype(str).str.upper().eq("RURAL").any()
        disc = [tc(d) for d in split_list(c["Discapacidades"])]
        fila = {
            "DANE institución": k, "Institución": tc(c["Nombre"]), "Sector": tc(c["Sector"]),
            "En el buscador": si_no(h),
            "Comuna": h["comuna"] if h else "", "Núcleo": h["nucleo"] if h else "",
            "Dirección": (h or {}).get("dir") or (c["Dirección"] if isinstance(c["Dirección"], str) else ""),
            "Teléfono fijo": (h or {}).get("tel") or " - ".join(fijos(c["Teléfono"])),
            "Correo institucional": (h or {}).get("correo") or " ".join(correos_inst(c["Correo electrónico"])),
            "Ciclos CLEI": clei_txt(gr),
        }
        for g in range(21, 27):
            fila[f"CLEI {g - 20}"] = si_no(g in gr)
        fila.update({
            "Modelo para adultos": ", ".join(mod_adultos.get(k, [])),
            "Jornadas": ", ".join(jor),
            "Fin de semana": si_no("Fin de semana" in jor),
            "Sedes con CLEI y fin de semana": len(sedes_fs),
            "Cuáles sedes (CLEI y fin de semana)": ", ".join(sedes_fs),
            "Educación para adultos": si_no(c["Atiende adultos"] == "Sí"),
            "Atiende discapacidad": si_no(disc),
            "Tipos de discapacidad": ", ".join(disc),
            "Talentos excepcionales": si_no(split_list(c["Capacidades excepcionales"])),
            "Sede en zona rural": si_no(rural),
            "Sedes activas": int(c["Sedes activas"]),
        })
        filas.append(fila)
    df = pd.DataFrame(filas).sort_values(["Fin de semana", "Institución"], ascending=[False, True])
    en = df[df["En el buscador"] == "Sí"].drop(columns=["En el buscador"])
    fuera = df[df["En el buscador"] == "No"].drop(columns=["En el buscador", "Comuna", "Núcleo"])
    df_sedes = pd.DataFrame(filas_sede)
    df_sedes = df_sedes[df_sedes["DANE institución"].isin(df["DANE institución"])]

    def resumen(d, nombre):
        n = len(d)
        pct = lambda x: f"{x * 100 / n:.0f}%" if n else ""
        r = [{"Grupo": nombre, "Indicador": "Instituciones con CLEI", "Instituciones": n, "%": "100%"}]
        for col in ("Fin de semana", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"):
            k = int((d[col] == "Sí").sum())
            r.append({"Grupo": nombre, "Indicador": col, "Instituciones": k, "%": pct(k)})
        k = int((d["Sedes con CLEI y fin de semana"] > 0).sum())
        r.append({"Grupo": nombre, "Indicador": "CLEI y fin de semana en la misma sede", "Instituciones": k, "%": pct(k)})
        return r
    res = pd.DataFrame(resumen(en, "En el buscador") + resumen(fuera, "Fuera del buscador (base SINEB)"))
    cruce = (en.assign(**{"Total": 1})
             .pivot_table(index="Fin de semana", values=["Total", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"],
                          aggfunc=lambda v: int(((v == "Sí") | (v == 1)).sum()))
             .reindex(columns=["Total", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"])
             .reset_index().rename(columns={"Fin de semana": "¿Tiene fin de semana?"}))
    notas = pd.DataFrame({"Notas": [
        "Instituciones con CLEI activo (Ciclos Lectivos Especiales Integrados: educación formal de jóvenes y adultos).",
        "Fuente: SINEB – MinEducación (Buscando Colegio), consulta 2026-10-16; comuna, núcleo y contacto del buscador.",
        "CLEI 1 = 1° a 3°, CLEI 2 = 4° y 5°, CLEI 3 = 6° y 7°, CLEI 4 = 8° y 9°, CLEI 5 = 10°, CLEI 6 = 11°.",
        "Las filas en amarillo tienen jornada de FIN DE SEMANA. 'CLEI y fin de semana en la misma sede' confirma que ambas se ofrecen en la misma sede.",
        "Atención e inclusión: Educación para adultos, Atiende discapacidad, Talentos excepcionales y Sede en zona rural (columnas Sí/No).",
        "La jornada nocturna no se incluye, por decisión del proyecto.",
        "Privacidad: solo teléfonos fijos y correos institucionales; sin nombres de rectores.",
        "La base no indica qué grado se dicta en qué jornada: confirme con la institución.",
    ]})
    path = OUT / "instituciones_clei.xlsx"
    write_xlsx(path, {"Leeme": notas, "Resumen": res, "Fin de semana x inclusión": cruce,
                      "CLEI en el buscador": en, "CLEI fuera del buscador": fuera, "Sedes": df_sedes})
    resaltar_fin_de_semana(path, ["CLEI en el buscador", "CLEI fuera del buscador", "Sedes"])
    print(f"datos/instituciones_clei.xlsx: {len(en)} en el buscador + {len(fuera)} fuera; "
          f"fin de semana {int((en['Fin de semana'] == 'Sí').sum())}/{len(en)}")


def resaltar_fin_de_semana(path, hojas):
    """Pinta en amarillo institucional las filas con jornada de fin de semana."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    amarillo = PatternFill("solid", fgColor="FFF4B8")
    for h in hojas:
        ws = wb[h]
        cab = [c.value for c in ws[1]]
        col = cab.index("Fin de semana") + 1
        for fila in ws.iter_rows(min_row=2):
            if fila[col - 1].value == "Sí":
                for c in fila:
                    c.fill = amarillo
    wb.save(path)


def sanear_fuente():
    """Reescribe la fuente sin datos personales (todas las hojas), conservando la estructura."""
    xl = pd.ExcelFile(SRC)
    hojas = {n: (xl.parse(n) if n == "Leeme" else sanear_df(xl.parse(n))) for n in xl.sheet_names}
    write_xlsx(SRC, hojas)
    print(f"{SRC.name}: fuente saneada ({', '.join(hojas)})")


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
    if "--sanear-fuente" in sys.argv:
        sanear_fuente()
    main()
