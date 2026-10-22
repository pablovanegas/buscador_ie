# Datos del buscador

| Archivo | Contenido |
|---|---|
| `fuente/colegios_medellin_sineb.xlsx` | Base de SINEB – MinEducación («Buscando Colegio»), Medellín, consulta del 2026-10-16. 568 instituciones y 759 sedes. **Copia saneada:** sin rectores, correos personales, celulares ni jornada nocturna (ver «Privacidad» y «Jornada nocturna»). |
| `fuente/IE_con_Modelos_Flexibles_2026_V2.xlsx` | Secretaría de Educación de Medellín: grupos de **modelos flexibles 2026** (Brújula, Aceleración del Aprendizaje, Pensar 1, 2 y 3), con institución, sede, comuna, jornada y estudiantes por grupo. **Define qué instituciones publica la página.** |
| `instituciones_no_incluidas.xlsx` | Las 338 instituciones de la base que **no** están en el buscador: 335 no oficiales (privadas) y 3 oficiales. Incluye sus sedes, modelos, grados y tarifas publicadas. |
| `instituciones_clei.xlsx` | Oferta para jóvenes y adultos, del buscador y de la base SINEB: **IES CLEI** (116) e **IES Modelos flexibles** (212, con una columna por modelo), más las hojas de CLEI en/fuera del buscador, un resumen por tema y un resumen por modelo. Distingue **fin de semana** (filas en amarillo), si CLEI y el fin de semana coinciden en la misma sede, y la **atención e inclusión** (adultos, discapacidad, talentos excepcionales, zona rural). |
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
5. **Jornada nocturna:** por decisión del proyecto, no se publica en ningún lado: ni en la página ni en los Excel. El script la quita de las listas de jornadas y elimina las columnas que la marcan, también en la fuente saneada.
6. **Lo que no se usó:** las columnas de tarifas, ICFES, computadores, internet y estrato vienen vacías o casi vacías para las instituciones oficiales, así que no se incluyeron.

## Modelos flexibles 2026 (página)

La página muestra **solo** las 104 instituciones del archivo de modelos flexibles 2026 y sus 472 grupos (11.557 estudiantes). El script lo lee así:

1. Toma las hojas de Brújula/Aceleración y de Pensar, quita las filas de «Total general» y normaliza el modelo (`BRUJULA ` con espacio final, `AA./BRUJ` = grupo mixto Aceleración + Brújula).
2. Cruza cada institución por nombre con el Consolidado de SINEB para obtener su DANE. Las 104 cruzan: 100 ya estaban en el buscador y 4 privadas no (se agregan en `MF_EXTRA` con datos de SINEB y sin coordenadas).
3. Escribe en `index.html` las variables `MF` (grupos por institución: sede, comuna, jornada, modelo, estudiantes), `MF_EXTRA` y `MF_META`.

La hoja «Cantidad de estudiantes» del archivo (11.548 estudiantes, 416 grupos) no cuadra con el detalle por grupo; se usa el detalle.

## Actualizar

Reemplaza el Excel de `fuente/` por una consulta nueva y ejecuta:

```bash
python3 scripts/integrar_sineb.py --sanear-fuente
```

`--sanear-fuente` quita los datos personales y la jornada nocturna de la consulta nueva antes de guardarla en el repositorio. No subas la consulta original sin sanear.

Para actualizar los modelos flexibles, reemplaza `fuente/IE_con_Modelos_Flexibles_2026_V2.xlsx` con la misma estructura de hojas y vuelve a ejecutar el script.

El script es idempotente: vuelve a generar las variables `SINEB`, `MF`, `MF_EXTRA` y `MF_META` en `index.html` y los libros de Excel.
