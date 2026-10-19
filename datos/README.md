# Datos del buscador

| Archivo | Contenido |
|---|---|
| `fuente/colegios_medellin_sineb.xlsx` | Base original de SINEB – MinEducación («Buscando Colegio»), Medellín, consulta del 2026-10-16. 568 instituciones y 759 sedes. |
| `instituciones_no_incluidas.xlsx` | Las 338 instituciones de la base que **no** están en el buscador: 335 no oficiales (privadas) y 3 oficiales. Incluye sus sedes, modelos, grados y tarifas publicadas. |
| `reporte_cruce_html_sineb.xlsx` | Calidad del cruce entre el buscador y SINEB, con las diferencias de nombre, rector, teléfono y correo de cada institución. |

## Cómo se integró

1. **Llave de cruce:** el buscador usa el código DANE de la sede principal (14 dígitos), que es el DANE de la institución (12 dígitos) seguido de `01`. Las 230 instituciones del buscador cruzan con SINEB.
2. **Qué se agregó a cada institución** (variable `SINEB` en `index.html`):
   - jornadas escolares;
   - grados ofrecidos, incluidos los ciclos para adultos (CLEI), Aceleración del aprendizaje y la formación complementaria de la Normal;
   - modelos educativos flexibles y si atiende adultos;
   - discapacidades atendidas, capacidades excepcionales, etnias e idiomas;
   - tipo, carácter, especialidad, género, zona y prestador del servicio;
   - todas las sedes activas, con su dirección, teléfono, jornadas, grados y modelos;
   - teléfonos, correos y rector según SINEB.
3. **Datos de contacto:** se conservan los datos de contacto del buscador. Cuando SINEB tiene otros, la ficha los muestra aparte, marcados como «según SINEB».
4. **Lo que no se usó:** las columnas de tarifas, ICFES, computadores, internet y estrato vienen vacías o casi vacías para las instituciones oficiales, así que no se incluyeron.

## Actualizar

Reemplaza el Excel de `fuente/` por una consulta nueva y ejecuta:

```bash
python3 scripts/integrar_sineb.py
```

El script es idempotente: vuelve a generar la variable `SINEB` en `index.html` y los dos libros de Excel.
