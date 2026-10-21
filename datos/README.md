# Datos del buscador

| Archivo | Contenido |
|---|---|
| `fuente/colegios_medellin_sineb.xlsx` | Base de SINEB – MinEducación («Buscando Colegio»), Medellín, consulta del 2026-10-16. 568 instituciones y 759 sedes. **Copia saneada:** sin rectores, correos personales ni celulares (ver «Privacidad»). |
| `instituciones_no_incluidas.xlsx` | Las 338 instituciones de la base que **no** están en el buscador: 335 no oficiales (privadas) y 3 oficiales. Incluye sus sedes, modelos, grados y tarifas publicadas. |
| `instituciones_clei.xlsx` | Oferta para jóvenes y adultos, del buscador y de la base SINEB: **IES CLEI** (116), **IES Nocturna** (77) e **IES Modelos flexibles** (212, con una columna por modelo), más las hojas de CLEI en/fuera del buscador, un resumen por tema y un resumen por modelo. Distingue **fin de semana** (filas en amarillo), **nocturna** (celdas en azul), si CLEI y la jornada coinciden en la misma sede, y la **atención e inclusión** (adultos, discapacidad, talentos excepcionales, zona rural). |
| `reporte_cruce_html_sineb.xlsx` | Calidad del cruce entre el buscador y SINEB, con las diferencias de nombre, teléfono y correo institucionales de cada institución. |

## Cómo se integró

1. **Llave de cruce:** el buscador usa el código DANE de la sede principal (14 dígitos), que es el DANE de la institución (12 dígitos) seguido de `01`. Las 230 instituciones del buscador cruzan con SINEB.
2. **Qué se agregó a cada institución** (variable `SINEB` en `index.html`):
   - jornadas escolares;
   - grados ofrecidos, incluidos los ciclos para adultos (CLEI), Aceleración del aprendizaje y la formación complementaria de la Normal;
   - modelos educativos flexibles, separando los dirigidos a adultos (población «ADULTOS» de la hoja Modelos), y si atiende adultos;
   - discapacidades atendidas, capacidades excepcionales, etnias e idiomas;
   - tipo, carácter, especialidad, género, zona y prestador del servicio;
   - todas las sedes activas, con su dirección, teléfono fijo, jornadas, grados y modelos;
   - teléfonos fijos y correos institucionales según SINEB.
3. **Privacidad:** no se publica ningún dato personal, ni en la página ni en los Excel. Se eliminan los nombres de rectores, los correos que no son de un dominio institucional (`.gov.co` o `.edu.co`), los correos institucionales con el nombre de una persona (lista `CORREOS_PERSONALES` del script) y los números de celular. Solo quedan teléfonos fijos y correos institucionales.
4. **Datos de contacto:** se conservan los datos de contacto del buscador. Cuando SINEB tiene otros, la ficha los muestra aparte, marcados como «según SINEB».
5. **Jornada nocturna:** por decisión del proyecto, la página no publica la jornada nocturna; el script la omite en la variable `SINEB`. El libro `instituciones_clei.xlsx` sí la incluye, para análisis interno, y los demás Excel conservan las columnas originales de la fuente.
6. **Lo que no se usó:** las columnas de tarifas, ICFES, computadores, internet y estrato vienen vacías o casi vacías para las instituciones oficiales, así que no se incluyeron.

## Actualizar

Reemplaza el Excel de `fuente/` por una consulta nueva y ejecuta:

```bash
python3 scripts/integrar_sineb.py --sanear-fuente
```

`--sanear-fuente` quita los datos personales de la consulta nueva antes de guardarla en el repositorio. No subas la consulta original sin sanear.

El script es idempotente: vuelve a generar la variable `SINEB` en `index.html` y los libros de Excel.
