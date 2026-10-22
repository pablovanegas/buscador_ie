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
# Grupos de modelos flexibles 2026 (Secretaría de Educación): define qué instituciones muestra la página
SRC_MF = ROOT / "datos/fuente/IE_con_Modelos_Flexibles_2026_V2.xlsx"
MF_COD = {"BRUJULA": "BRU", "ACELERACIÓN": "ACE", "AA./BRUJ": "MIX", "PENSAR 1": "P1", "PENSAR 2": "P2", "PENSAR 3": "P3"}
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


def sin_nocturna(v):
    """Quita NOCTURNA de una lista de jornadas, conservando el separador original."""
    if not isinstance(v, str):
        return v
    sep = ", " if ", " in v else ","
    return sep.join(p for p in split_list(v) if p.upper() != "NOCTURNA") or None


def sanear_df(df):
    """Quita rectores, correos personales, celulares y la jornada nocturna de cualquier hoja de la base."""
    df = df.drop(columns=[c for c in df.columns if "rector" in c.lower() or "nocturn" in c.lower()])
    for c in df.columns:
        if c.lower().startswith("jornada"):
            df[c] = df[c].map(sin_nocturna)
        elif c.lower().startswith("correo"):
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
# La jornada NOCTURNA se excluye a propósito: el proyecto no publica esa información (ni en la página ni en los Excel).
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


def cargar_modelos_flexibles(cons):
    """Lee el archivo de grupos 2026 (hojas Brújula/Aceleración y Pensar) y cruza cada institución
    con SINEB por nombre para obtener su código DANE. Devuelve {dane12: [[sede, comuna, jornada, modelo, estudiantes], ...]}."""
    xl = pd.ExcelFile(SRC_MF)
    hojas = [h for h in xl.sheet_names if "BRUJULA" in h.upper() or "PENSAR" in h.upper()]
    df = pd.concat([xl.parse(h) for h in hojas], ignore_index=True)
    df = df[df["SEDE"].notna() & df["GRUPO"].notna()]           # quita la fila "Total general"
    df["COD"] = df["GRUPO"].astype(str).str.strip().str.upper().map(MF_COD)
    assert df["COD"].notna().all(), f"Grupos sin código: {sorted(df[df['COD'].isna()]['GRUPO'].unique())}"
    por_nombre = {norm(n): str(k) for n, k in zip(cons["Nombre"], cons["Código DANE"])}
    df["_k"] = df["INSTITUCION"].map(lambda n: por_nombre.get(norm(n)))
    sin = sorted(df[df["_k"].isna()]["INSTITUCION"].unique())
    assert not sin, f"Instituciones del archivo sin código DANE en SINEB: {sin}"
    mf = {}
    for _, r in df.iterrows():
        mf.setdefault(r["_k"], []).append([tc(r["SEDE"]), int(r["COMUNA"]), tc(r["JORNADA"]), r["COD"], int(r["Total"])])
    print(f"Modelos flexibles 2026: {len(mf)} instituciones · {df['SEDE'].nunique()} sedes · {len(df)} grupos · {int(df['Total'].sum())} estudiantes")
    return mf


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

    # Modelos flexibles 2026: instituciones del archivo que no están en el buscador (privadas) se agregan
    # con los datos de SINEB, sin coordenadas.
    mf12 = cargar_modelos_flexibles(cons) if SRC_MF.exists() else {}
    comuna_lbl = {r["comuna"].split(" - ")[0]: r["comuna"] for r in data if r.get("comuna")}
    extras, dkey = [], {k: html_by_key[k]["dane"] for k in matched}
    for k in sorted(set(mf12) - set(html_by_key)):
        c = cons_k.loc[k]
        num = str(mf12[k][0][1])
        dkey[k] = k + "01"
        extras.append({
            "nombre": str(c["Nombre"]), "dane": k + "01", "nit": "", "nucleo": "",
            "tel": " - ".join(fijos(c["Teléfono"])), "correo": " ".join(correos_inst(c["Correo electrónico"])),
            "dir": c["Dirección"] if isinstance(c["Dirección"], str) else "", "barrio": "",
            "comuna": comuna_lbl.get(num, num), "rector": None,
            "trans": False, "trans_t": "", "mt": False, "jc": False, "pae": False, "prg_mt": [], "prg_jc": [],
            "lat": None, "lon": None, "privada": c["Sector"] == "NO OFICIAL",
        })
        extras[-1].pop("rector")

    # ── registro compacto por institución ───────────────────────────
    sineb = {}
    for k in matched + [e["dane"][:12] for e in extras]:
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
        sineb[dkey[k]] = {a: b for a, b in rec.items() if b not in ([], "", None)}

    # ── inserta / reemplaza la línea SINEB en el HTML ───────────────
    leeme = " ".join(str(v) for v in xl.parse("Leeme").iloc[:, 0])
    gen = re.search(r"Generado:\s*(\d{4}-\d{2}-\d{2})", leeme)
    meta = {"fuente": "SINEB – MinEducación (Buscando Colegio)", "fecha": gen.group(1) if gen else ""}
    line = "var SINEB=" + json.dumps(sineb, ensure_ascii=False, separators=(",", ":")) + ";"
    meta_line = "var SINEB_META=" + json.dumps(meta, ensure_ascii=False) + ";"
    lines = html.split("\n")
    lines = [l for l in lines if not l.startswith(("var SINEB=", "var SINEB_META=", "var MF=", "var MF_EXTRA=", "var MF_META="))]
    i = next(n for n, l in enumerate(lines) if l.startswith("var DATA="))
    lines[i] = "var DATA=" + json.dumps(data, ensure_ascii=False) + ";"
    mf = {dkey[k]: v for k, v in mf12.items()}
    est = sum(r[4] for v in mf.values() for r in v)
    mf_meta = {"fuente": "Secretaría de Educación de Medellín – IE con modelos flexibles 2026 (V2)",
               "instituciones": len(mf), "grupos": sum(len(v) for v in mf.values()), "estudiantes": est}
    lines[i + 1:i + 1] = [line, meta_line,
                          "var MF=" + json.dumps(mf, ensure_ascii=False, separators=(",", ":")) + ";",
                          "var MF_EXTRA=" + json.dumps(extras, ensure_ascii=False, separators=(",", ":")) + ";",
                          "var MF_META=" + json.dumps(mf_meta, ensure_ascii=False) + ";"]
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

    exportar_oferta(cons, sedes, html_by_key, mods)


# ── Excel 3: oferta para jóvenes y adultos (CLEI y modelos flexibles) ──
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


def exportar_oferta(cons, sedes, html_by_key, mods):
    """Libro instituciones_clei.xlsx: una fila por institución (buscador y base SINEB) con CLEI,
    fin de semana, modelos flexibles e inclusión. Solo contactos institucionales."""
    si_no = lambda b: "Sí" if b else "No"
    # Modelos flexibles por institución y su población
    pobl = {tc(m): ("Adultos" if p == "ADULTOS" else "Niños y jóvenes") for m, p in zip(mods["Modelo"], mods["Población"])}
    mod_inst = {str(k): sorted({tc(m) for m in g["Modelo"]}) for k, g in mods.groupby("Código DANE")}
    modelos = sorted(pobl, key=lambda m: (pobl[m] != "Niños y jóvenes", m))

    filas = []
    for _, c in cons.iterrows():
        k = c["_k"]
        h = html_by_key.get(k)
        gr = grades(c["Grados (códigos)"])
        jor = jornadas(c["Jornadas (ficha)"])
        sd = sedes[sedes["_k"] == k]
        s_fs = []
        for _, x in sd.iterrows():
            xgr, xjor = grades(x["Grados (códigos)"]), jornadas(x["Jornadas"])
            if cleis(xgr) and "Fin de semana" in xjor:
                s_fs.append(tc(x["Nombre sede"]))
        mf = mod_inst.get(k, [])
        disc = [tc(d) for d in split_list(c["Discapacidades"])]
        rural = "RURAL" in str(c["Zona"]).upper() or sd["Zona"].astype(str).str.upper().eq("RURAL").any()
        f = {
            "DANE institución": k, "Institución": tc(c["Nombre"]), "Sector": tc(c["Sector"]),
            "En el buscador": si_no(h), "Comuna": h["comuna"] if h else "", "Núcleo": h["nucleo"] if h else "",
            "Dirección": (h or {}).get("dir") or (c["Dirección"] if isinstance(c["Dirección"], str) else ""),
            "Teléfono fijo": (h or {}).get("tel") or " - ".join(fijos(c["Teléfono"])),
            "Correo institucional": (h or {}).get("correo") or " ".join(correos_inst(c["Correo electrónico"])),
            "Jornadas": ", ".join(jor),
            "Fin de semana": si_no("Fin de semana" in jor),
            "Ciclos CLEI": clei_txt(gr),
        }
        for g in range(21, 27):
            f[f"CLEI {g - 20}"] = si_no(g in gr)
        f.update({
            "Sedes con CLEI y fin de semana": len(s_fs), "Cuáles sedes (CLEI y fin de semana)": ", ".join(s_fs),
            "N.º modelos flexibles": len(mf),
            "Modelos para niños y jóvenes": ", ".join(m for m in mf if pobl.get(m) == "Niños y jóvenes"),
            "Modelos para adultos": ", ".join(m for m in mf if pobl.get(m) == "Adultos"),
            "Educación para adultos": si_no(c["Atiende adultos"] == "Sí"),
            "Atiende discapacidad": si_no(disc), "Tipos de discapacidad": ", ".join(disc),
            "Talentos excepcionales": si_no(split_list(c["Capacidades excepcionales"])),
            "Sede en zona rural": si_no(rural), "Sedes activas": int(c["Sedes activas"]),
        })
        for m in modelos:
            f[m] = si_no(m in mf)
        f["_clei"], f["_mf"] = bool(cleis(gr)), bool(mf)
        filas.append(f)

    todo = pd.DataFrame(filas).sort_values(["En el buscador", "Institución"], ascending=[False, True])
    base = [c for c in todo.columns if c not in modelos and not c.startswith("_")]
    sin = lambda d, cols: d.drop(columns=[c for c in cols if c in d.columns])
    clei = todo[todo["_clei"]]
    flex = todo[todo["_mf"]]

    # Hojas de CLEI
    cols_clei = base
    clei_en = sin(clei[clei["En el buscador"] == "Sí"][cols_clei], ["En el buscador"])
    clei_fu = sin(clei[clei["En el buscador"] == "No"][cols_clei], ["En el buscador", "Comuna", "Núcleo"])
    # Modelos: todas las instituciones (buscador y base SINEB) en una sola hoja
    ies_clei = clei[cols_clei]
    ies_flex = flex[[c for c in base if not c.startswith("Sedes con") and not c.startswith("Cuáles")] + modelos]

    # Resumen
    def bloque(d, tema):
        r = []
        for grupo, dd in (("En el buscador", d[d["En el buscador"] == "Sí"]), ("Fuera del buscador (base SINEB)", d[d["En el buscador"] == "No"])):
            n = len(dd)
            pct = lambda x: f"{x * 100 / n:.0f}%" if n else ""
            r.append({"Tema": tema, "Grupo": grupo, "Indicador": "Instituciones", "Instituciones": n, "%": "100%" if n else ""})
            for col in ("Fin de semana", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"):
                v = int((dd[col] == "Sí").sum())
                r.append({"Tema": tema, "Grupo": grupo, "Indicador": col, "Instituciones": v, "%": pct(v)})
            v = int((dd["N.º modelos flexibles"] > 0).sum())
            r.append({"Tema": tema, "Grupo": grupo, "Indicador": "Con algún modelo flexible", "Instituciones": v, "%": pct(v)})
            v = int(dd["_clei"].sum())
            r.append({"Tema": tema, "Grupo": grupo, "Indicador": "Con CLEI", "Instituciones": v, "%": pct(v)})
        return r
    res = pd.DataFrame(bloque(clei, "CLEI") + bloque(flex, "Modelos flexibles"))
    res = res[~((res["Tema"] == "CLEI") & (res["Indicador"] == "Con CLEI"))
              & ~((res["Tema"] == "Modelos flexibles") & (res["Indicador"] == "Con algún modelo flexible"))]
    mod_res = pd.DataFrame([{
        "Modelo flexible": m, "Población": pobl[m],
        "Instituciones en el buscador": int(((todo[m] == "Sí") & (todo["En el buscador"] == "Sí")).sum()),
        "Instituciones fuera del buscador": int(((todo[m] == "Sí") & (todo["En el buscador"] == "No")).sum()),
        "Con CLEI": int(((todo[m] == "Sí") & todo["_clei"]).sum()),
    } for m in modelos])
    en = clei_en
    cruce = (en.assign(Total=1)
             .pivot_table(index="Fin de semana", values=["Total", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"],
                          aggfunc=lambda v: int(((v == "Sí") | (v == 1)).sum()))
             .reindex(columns=["Total", "Educación para adultos", "Atiende discapacidad", "Talentos excepcionales", "Sede en zona rural"])
             .reset_index().rename(columns={"Fin de semana": "¿Tiene fin de semana?"}))
    notas = pd.DataFrame({"Notas": [
        "Oferta para jóvenes y adultos: instituciones con CLEI y con modelos educativos flexibles.",
        "Fuente: SINEB – MinEducación (Buscando Colegio), consulta 2026-10-16; comuna, núcleo y contacto del buscador.",
        "Hojas 'IES …': todas las instituciones del tema, del buscador y de la base SINEB (columna 'En el buscador').",
        "CLEI 1 = 1° a 3°, CLEI 2 = 4° y 5°, CLEI 3 = 6° y 7°, CLEI 4 = 8° y 9°, CLEI 5 = 10°, CLEI 6 = 11°.",
        "Filas en amarillo: jornada de FIN DE SEMANA.",
        "'Sedes con CLEI y fin de semana' confirma que ambos se ofrecen en la misma sede.",
        "En 'IES Modelos flexibles' cada modelo tiene su columna (Sí/No); 'Modelos flexibles' resume cada modelo.",
        "Privacidad: solo teléfonos fijos y correos institucionales; sin nombres de rectores.",
        "La base no indica qué grado se dicta en qué jornada: confirme con la institución.",
    ]})
    path = OUT / "instituciones_clei.xlsx"
    hojas = {"Leeme": notas, "Resumen": res, "Fin de semana x inclusión": cruce,
             "IES CLEI": ies_clei, "CLEI en el buscador": clei_en, "CLEI fuera del buscador": clei_fu,
             "IES Modelos flexibles": ies_flex, "Modelos flexibles": mod_res}
    write_xlsx(path, hojas)
    resaltar(path, ["IES CLEI", "CLEI en el buscador", "CLEI fuera del buscador", "IES Modelos flexibles"])
    print(f"datos/instituciones_clei.xlsx: CLEI {len(clei)} · modelos flexibles {len(flex)} "
          f"({len(modelos)} modelos)")


def resaltar(path, hojas):
    """Filas con fin de semana en amarillo institucional."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    amarillo = PatternFill("solid", fgColor="FFF4B8")
    for h in hojas:
        ws = wb[h]
        cab = [c.value for c in ws[1]]
        fs = cab.index("Fin de semana")
        for fila in ws.iter_rows(min_row=2):
            if fila[fs].value == "Sí":
                for c in fila:
                    c.fill = amarillo
    wb.save(path)


def sanear_fuente():
    """Reescribe la fuente sin datos personales ni jornada nocturna (todas las hojas), conservando la estructura."""
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
