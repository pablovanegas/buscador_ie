# Instituciones con Modelos Flexibles 2026 · Medellín

Buscador de instituciones educativas de la Secretaría de Educación de Medellín. Incluye todas las instituciones del buscador y permite filtrar las que ofrecen **modelos flexibles** en 2026: **Brújula**, **Aceleración del Aprendizaje** y **Pensar 1, 2 y 3**.

### 👉 [Abrir el buscador](https://pablovanegas.github.io/buscador_ie/)

**230 instituciones oficiales** · de ellas, **100 con modelos flexibles 2026**

## Enlaces útiles

| Enlace | Para qué sirve |
|---|---|
| [Busca tu cupo](https://www.medellin.edu.co/tucupo/) | Solicitar o consultar un cupo escolar en una institución oficial de Medellín. |
| [Buscador de colegios a nivel nacional](https://sineb.mineducacion.gov.co/bcol/app?service=page/BuscandoColegioBasico) | Consultar cualquier institución del país (MinEducación, SINEB): sedes, jornadas y grados. |
| [Tableros Power BI](https://app.powerbi.com/view?r=eyJrIjoiZThmMjY3ZjQtZGE3Ni00YmU4LThlMzUtMTk3MTEzNTlmZDhjIiwidCI6IjljNDhlMDg4LTVlNDQtNGIwZC05M2EwLWVlYjJjNjEyN2MzZCIsImMiOjR9&embedImagePlaceholder=true) | Tableros de datos para consulta y análisis. |

Los tres enlaces también están en la parte superior del buscador.

## Guía rápida

1. **Busca** en la caja grande por nombre de la institución, sede, barrio o código DANE. La tecla `/` lleva directo a ella.
2. **Filtra** debajo de la caja de búsqueda por comuna o corregimiento, núcleo educativo, modelo flexible (menú desplegable, con «Con algún modelo flexible») y jornada. El panel lateral suma grado, tipo y especialidad y servicios.
3. **Elige la vista**:
   - **Tarjetas**: resumen de cada institución; las que tienen modelos flexibles muestran cuáles ofrecen y en qué jornadas.
   - **Mapa**: ubicación de la sede principal de cada institución.
   - **Guía**: qué es cada modelo y cómo leer los datos.
4. **Ficha**: abre el detalle de la institución: contacto, oferta educativa (modelos flexibles con sus jornadas y sedes, jornadas y grados regulares), sedes y servicios.
5. **Busca tu cupo**: el botón amarillo está en cada tarjeta, ficha y punto del mapa.

### Los modelos

| Modelo | Grados que cubre en un año |
|---|---|
| Brújula | 1°, 2° y 3° |
| Aceleración del Aprendizaje | 4° y 5° |
| Pensar 1 | 6° y 7° |
| Pensar 2 | 8° y 9° |
| Pensar 3 | 10° y 11° |

Son para estudiantes en **extraedad** (dos o más años por encima de la edad de su grado).

## Datos

- **Datos del buscador**: contacto, ubicación y servicios de las instituciones.
- **Modelos flexibles 2026**, Secretaría de Educación de Medellín: qué modelos ofrece cada una de las 100 instituciones oficiales, en qué sedes y jornadas.
- **SINEB – MinEducación**: jornadas, grados y sedes de cada institución.
- Solo se publican **instituciones oficiales**: las privadas no aparecen.
- Solo se publican **contactos institucionales**: no hay nombres de rectores, correos personales ni celulares.
- No se publican la **jornada nocturna** (ni en la página ni en los Excel) ni la **jornada completa**, ni cifras de grupos o estudiantes.

El detalle de las fuentes, el cruce y cómo actualizar está en [`datos/README.md`](datos/README.md). Para regenerar los datos de la página:

```bash
python3 scripts/integrar_sineb.py
```

## Estructura

| Ruta | Contenido |
|---|---|
| `index.html` | El buscador completo (una sola página, con los datos incluidos). |
| `assets/` | Logo de la Alcaldía de Medellín. |
| `datos/` | Fuentes en Excel y libros de apoyo. |
| `scripts/integrar_sineb.py` | Genera los datos de la página y los libros de Excel. |
