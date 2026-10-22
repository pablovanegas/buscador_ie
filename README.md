# Instituciones con Modelos Flexibles 2026 · Medellín

Buscador de la Secretaría de Educación de Medellín con las instituciones educativas que tienen grupos de **modelos flexibles** en 2026: **Brújula**, **Aceleración del Aprendizaje** y **Pensar 1, 2 y 3**.

### 👉 [Abrir el buscador](https://pablovanegas.github.io/buscador_ie/)

**104 instituciones · 142 sedes · 472 grupos · 11.557 estudiantes**

## Enlaces útiles

| Enlace | Para qué sirve |
|---|---|
| [Busca tu cupo](https://www.medellin.edu.co/tucupo/) | Solicitar o consultar un cupo escolar en una institución oficial de Medellín. |
| [Buscador de colegios a nivel nacional](https://sineb.mineducacion.gov.co/bcol/app?service=page/BuscandoColegioBasico) | Consultar cualquier institución del país (MinEducación, SINEB): sedes, jornadas y grados. |
| [Tableros Power BI](https://app.powerbi.com/view?r=eyJrIjoiZThmMjY3ZjQtZGE3Ni00YmU4LThlMzUtMTk3MTEzNTlmZDhjIiwidCI6IjljNDhlMDg4LTVlNDQtNGIwZC05M2EwLWVlYjJjNjEyN2MzZCIsImMiOjR9&embedImagePlaceholder=true) | Tableros de datos para consulta y análisis. |

Los tres enlaces también están en la franja azul de la parte superior del buscador.

## Guía rápida

1. **Busca** en la caja grande por nombre de la institución, sede, barrio o código DANE. La tecla `/` lleva directo a ella.
2. **Filtra** por modelo, jornada del grupo y comuna en «Encuentra un grupo de modelo flexible», o con los botones de modelo del panel lateral.
3. **Elige la vista**:
   - **Tarjetas**: resumen de cada institución con sus grupos y estudiantes.
   - **Tabla**: un grupo por fila, ordenable. **Descargar CSV** baja lo que estás viendo (abre en Excel).
   - **Gráficos**: estudiantes por modelo, jornada y comuna. Un clic en una barra filtra.
   - **Mapa**: ubicación de la sede principal. 4 instituciones privadas no tienen coordenadas.
   - **Guía**: qué es cada modelo y cómo leer los datos.
4. **Ficha**: abre el detalle de la institución con los grupos por sede, contacto, oferta, sedes y servicios.

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

- **Modelos flexibles 2026**, Secretaría de Educación de Medellín: define qué instituciones aparecen y aporta los grupos, las jornadas y los estudiantes.
- **SINEB – MinEducación**: jornadas, grados, sedes e inclusión de cada institución.
- Solo se publican **contactos institucionales**: no hay nombres de rectores, correos personales ni celulares.

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
