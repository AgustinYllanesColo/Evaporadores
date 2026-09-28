---
name: "presupuesto-refrivan"
description: "Completar la plantilla PRESUPUESTO REFRIVAN.xlsx (modelo nuevo, con logo y colores de la marca) a partir de un párrafo con los servicios y el precio, y guardar Excel + PDF en la carpeta Presupuestos. Usar cuando Agustín pida armar un presupuesto."
---

# Presupuesto REFRIVAN

Entrada: la plantilla `PRESUPUESTO REFRIVAN.xlsx` (hoja "Presupuesto") + un párrafo informal con cliente, trabajos/productos y precio (ej: "cliente Marcelo, 4 heladeras verticales de 5 puertas, 5.000.000 cada una +iva").

Salida: DOS archivos, `Presupuesto N<numero>.xlsx` y `Presupuesto N<numero>.pdf`, completando la plantilla (nunca rehacerla de cero; conservar logo, franja tricolor, colores, bordes, formatos y fórmulas). El PDF es lo que se le manda al cliente.

## Ubicaciones

- Plantilla: viene dentro de esta skill, en `assets/PRESUPUESTO REFRIVAN.xlsx`. Si Agustín tiene una copia en `C:\Users\Usuario\OneDrive\Desktop\PLANTILLAS\PRESUPUESTO REFRIVAN.xlsx`, vale esa. La vieja `PRESUPUESTO LIMPIO.xlsx` ya NO se usa.
- Destino del Excel y el PDF: `C:\Users\Usuario\OneDrive\Desktop\Presupuestos\`
- Si la carpeta no está conectada a la sesión, pedir acceso en un solo pedido.
- Si alguna vez hay que regenerar la plantilla (cambio de dirección, teléfono, logo): `scripts/build_template.py assets/logo.png "assets/PRESUPUESTO REFRIVAN.xlsx"`.

## Paso 0: chequear los datos obligatorios ANTES de armar nada

Antes de tocar la plantilla, revisar el párrafo. Si falta algo de esta lista, PREGUNTAR y esperar la respuesta. No inventar ni asumir.

1. **Cliente**: tiene que venir al menos un dato del cliente (con el nombre alcanza, ej: "cliente Marcelo"). Si no dice nada del cliente, preguntar "¿A nombre de quién va el presupuesto?". El resto (contacto, dirección, teléfono, localidad, provincia) es opcional: si no viene, queda vacío sin preguntar.
2. **Ítems**: mínimo un ítem (trabajo o producto). Si no hay, preguntar qué se cotiza.
3. **Precio**: mínimo un precio (por ítem o global). Si no hay, preguntar el precio. Si el precio viene en USD, pasarlo a pesos con el dólar BNA del día (regla 3) y redondear al peso para arriba.
4. **IVA**: si el párrafo no lo aclara, preguntar: "¿Va sin IVA o más IVA? Si es más IVA, ¿21% o 10,5%?". Nunca asumir IVA 0. Que un costo diga "más IVA" (el IVA de lo que se compró) NO define el IVA del presupuesto: preguntar igual.

Preguntar en texto normal, corto, una cosa por vez y en ese orden. Nada de tarjetas de opción múltiple. Recién con los 4 puntos resueltos, armar el presupuesto.

## Mapa de celdas

- H3 = número del presupuesto (solo el número, ej. `129`; la etiqueta "N°" ya está en G3) · H4 = Fecha
- B10 Cliente (nombre de la empresa o persona) · G10 Contacto · B11 Dirección · G11 Localidad · B12 Teléfono · G12 Provincia
- Ítems: filas 16 a 24 → A = N°, B:D (combinadas) = Descripción, E = Modelo, F = Cant., G = Precio Unit., H = Subtotal
- H26 Subtotal (=SUM(H16:H24)) · F27/H27 IVA · H28 TOTAL (=H26+H27)
- D31 Cotización USD BNA · H31 Condición de pago
- A34 Observaciones

## Estilo del modelo (ya viene en la plantilla: no tocarlo)

Arial en todo. Colores de la marca: texto #283137, gris de etiquetas #5B6670, líneas #D5DBE0, recuadros #F1F4F6, franja tricolor rojo #BF222D / naranja #EB9E06 / verde #19988A. Encabezado de la tabla en #283137 con letra blanca. Filas de ítems blancas, solo con una línea fina abajo (sin franjas alternadas ni bordes verticales). Importes con `[$-2C0A]"$ "#,##0`, alineados a la derecha con sangría 1. Subtotales de la columna H en negrita.

Al cargar datos, escribir solo el VALOR de la celda: no cambiar fuentes, rellenos ni bordes, salvo lo que pide la regla 6.

## Reglas

1. **Número**: leer en memoria `/areas/presupuestos-refrivan.md` (y la copia del proyecto si existe; vale la más reciente) el último número emitido y usar el siguiente. Después de generar, actualizar esa línea. Si no hay registro, preguntar el número.
2. **Fecha**: fecha de hoy como valor fijo (un `date`, no `=TODAY()`), la celda ya tiene formato `dd/mm/yyyy`.
3. **Cotización USD BNA**: SIEMPRE consultar el dólar billete venta del Banco Nación del día (bna.com.ar) y escribirlo en D31 como texto con el formato `$ 1.535`. Nunca dejar el `$ 0` de la plantilla.
4. **Datos del cliente**: cargar solo lo que venga en el párrafo (o en la respuesta del Paso 0). Lo opcional que no venga queda vacío.
5. **Ítems**: un ítem por cada trabajo o producto que se vende por separado, numerados 1, 2, 3… Descripción corta que entre en una sola línea (sin ajuste de texto), con mayúscula inicial, sin la cantidad dentro del texto ("Limpieza de condensador y regulación de puertas"). La cantidad va en Cant. Nunca meter un párrafo largo en una sola celda.
6. **Producto con componentes** (kit, unidad armada, equipo con accesorios): es UN SOLO ítem con UN SOLO N°. Los componentes NO suman N°. Formato:
   - Primera fila: N°, nombre del producto en negrita (Arial 10 negrita, #283137), Modelo si corresponde (ej. "1 HP"), Cant., Precio Unit. y la fórmula del Subtotal.
   - Fila siguiente, sin N°, sin cantidad ni precio: "Incluye:" en Arial 9 negrita, #5B6670, sangría 1.
   - Una fila por componente, sin N°, sin cantidad ni precio: "•  Componente" en Arial 9, #5B6670, sangría 2.
   - En las filas de "Incluye:" y componentes, borrar la fórmula de H (dejarla vacía) y poner alto de fila 17.
   - Todo el bloque va unido: sacar la línea de abajo de la primera fila y de "Incluye:" y de cada componente, salvo el último, que conserva la línea fina #D5DBE0 abajo.
   - Si una parte la aporta el cliente (ej. el compresor), el trabajo va como componente ("Armado y conexión con el compresor del cliente") y se aclara en Observaciones (regla 13).
7. **Modelo**: vacío por defecto. Solo se completa si el párrafo dice un tipo, potencia o modelo específico (vitrina, cámara frigorífica, exhibidora, "1 HP", un modelo concreto). No poner "Heladera" genérico.
8. **Precio global**: si el precio es uno solo por todo el trabajo y hay varios ítems distintos, NO agregar un ítem extra de "mano de obra total". Escribir el importe directamente en el **Subtotal (columna H) del primer ítem**, dejando vacío su Precio Unit. (así la cantidad no multiplica el total). Los demás ítems quedan sin precio. Si es un solo producto con componentes (regla 6), el precio va en Precio Unit. de la primera fila, con su Cant. y la fórmula del subtotal. Si el párrafo da precios por ítem (o "cada uno"), cargarlos en Precio Unit. y dejar la fórmula del subtotal.
9. **IVA** (según el párrafo o la respuesta del Paso 0):
   - Sin IVA: F27 = `IVA  (0%)`, H27 = 0.
   - Más IVA 21%: F27 = `IVA  (21%)`, H27 = `=H26*0.21`.
   - Más IVA 10,5%: F27 = `IVA  (10,5%)`, H27 = `=H26*0.105`.
10. **Fórmulas de subtotal**: la plantilla ya trae `=IF(Fn*Gn=0,"",Fn*Gn)` en H16:H24 y B:D combinadas. Solo se borran en las filas de la regla 6 o se reemplazan por el importe en la regla 8.
11. **Más de 9 filas**: si los ítems (con sus componentes) no entran en las filas 16 a 24, avisar a Agustín y proponer resumir componentes; no insertar filas.
12. **Condición de pago**: SIEMPRE "100% contado".
13. **Observaciones** (A34): vacías, salvo que Agustín diga qué poner o que el cliente aporte una parte del equipo o del trabajo: eso SIEMPRE se aclara (ej. "El compresor lo provee el cliente.").

## Procedimiento

1. Paso 0: chequear cliente, ítems, precio e IVA; preguntar lo que falte.
2. Abrir la plantilla con openpyxl (conserva el logo).
3. Aplicar las reglas, guardar como `Presupuesto N<numero>.xlsx`.
4. Generar `Presupuesto N<numero>.pdf` con LibreOffice (`soffice --headless --convert-to pdf`), pasarlo a PNG y MIRARLO antes de entregar. Auditar: logo y franja tricolor visibles, número, fecha, cotización, un solo N° por producto, ninguna celda con un párrafo largo, subtotal/IVA/total, miles con punto, importes sin tocar el borde, observaciones y que entre en una sola hoja A4. Si algo se ve desprolijo, corregirlo antes de mostrarlo.
5. Guardar los DOS archivos (xlsx y pdf) en la carpeta Presupuestos de la PC.
6. Resumir en pocas líneas qué se cargó y qué quedó vacío.
7. Actualizar el contador en memoria.

Hablarle en castellano rioplatense, corto. Nada de tarjetas de opción múltiple: preguntar en texto normal, una cosa por vez.
