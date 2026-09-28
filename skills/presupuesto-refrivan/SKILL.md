---
name: "presupuesto-refrivan"
description: "Armar un presupuesto REFRIVAN (modelo nuevo con logo y colores de la marca, una o más hojas A4) a partir de un párrafo con los servicios y el precio, y guardar Excel + PDF en la carpeta Presupuestos. Usar cuando Agustín pida armar un presupuesto."
---

# Presupuesto REFRIVAN

Entrada: un párrafo informal con cliente, trabajos/productos y precio (ej: "cliente Marcelo, 4 heladeras verticales de 5 puertas, 5.000.000 cada una +iva").

Salida: DOS archivos, `Presupuesto N<numero>.xlsx` y `Presupuesto N<numero>.pdf`. El PDF es lo que se le manda al cliente.

El presupuesto lo arma SIEMPRE el script `scripts/generar_presupuesto.py` de esta skill: nunca editar celdas a mano ni usar la plantilla vieja `PRESUPUESTO LIMPIO.xlsx`. El script trae el diseño (logo, franja tricolor, colores, tipografía), calcula fórmulas y reparte en hojas.

## Ubicaciones

- Script: `scripts/generar_presupuesto.py` · Logo: `assets/logo.png` (los dos dentro de esta skill).
- Destino del Excel y el PDF: `C:\Users\Usuario\OneDrive\Desktop\Presupuestos\`
- Si esa carpeta no está conectada a la sesión, pedir acceso en un solo pedido.
- Necesita Python con `openpyxl` y `pymupdf`, y LibreOffice (`soffice`). Si falta algo, instalarlo (`pip install openpyxl pymupdf`).

## Paso 0: chequear los datos obligatorios ANTES de armar nada

Antes de armar, revisar el párrafo. Si falta algo de esta lista, PREGUNTAR y esperar la respuesta. No inventar ni asumir.

1. **Cliente**: tiene que venir al menos un dato del cliente (con el nombre alcanza, ej: "cliente Marcelo"). Si no dice nada del cliente, preguntar "¿A nombre de quién va el presupuesto?". El resto (contacto, dirección, teléfono, localidad, provincia) es opcional: si no viene, queda vacío sin preguntar.
2. **Ítems**: mínimo un ítem (trabajo o producto). Si no hay, preguntar qué se cotiza.
3. **Precio**: mínimo un precio (por ítem o global). Si no hay, preguntar el precio. Si el precio viene en USD, pasarlo a pesos con el dólar BNA del día (regla 3) y redondear al peso para arriba.
4. **IVA**: si el párrafo no lo aclara, preguntar: "¿Va sin IVA o más IVA? Si es más IVA, ¿21% o 10,5%?". Nunca asumir IVA 0. Que un costo diga "más IVA" (el IVA de lo que se compró) NO define el IVA del presupuesto: preguntar igual.

Preguntar en texto normal, corto, una cosa por vez y en ese orden. Nada de tarjetas de opción múltiple. Recién con los 4 puntos resueltos, armar el presupuesto.

## El JSON que recibe el script

```json
{
  "numero": 129,
  "fecha": "28/09/2026",
  "cliente": {"nombre": "Marcelo Gómez", "contacto": "", "telefono": "",
              "direccion": "", "localidad": "", "provincia": ""},
  "items": [
    {"descripcion": "Heladera vertical de 5 puertas", "modelo": "Exhibidora",
     "cantidad": 4, "precio_unitario": 5000000},
    {"descripcion": "Unidad condensadora sin compresor", "modelo": "1 HP",
     "cantidad": 1, "precio_unitario": 1850000,
     "incluye": ["Condensador con forzador", "Armado y conexión con el compresor del cliente"]}
  ],
  "precio_global": null,
  "iva": 21,
  "cotizacion_usd": 1535,
  "condicion_pago": "100% contado",
  "observaciones": "El compresor lo provee el cliente."
}
```

## Reglas

1. **Número**: leer en memoria `/areas/presupuestos-refrivan.md` (y la copia del proyecto si existe; vale la más reciente) el último número emitido y usar el siguiente. Después de generar, actualizar esa línea. Si no hay registro, preguntar el número.
2. **Fecha**: la de hoy, `dd/mm/yyyy`.
3. **Cotización USD BNA**: SIEMPRE consultar el dólar billete venta del Banco Nación del día (bna.com.ar) y ponerlo en `cotizacion_usd` como número (1535). Nunca reusar un valor viejo.
4. **Cliente**: cargar solo lo que venga en el párrafo (o en la respuesta del Paso 0). Lo que no venga va como `""`.
5. **Ítems**: un ítem por cada trabajo o producto que se vende por separado; el script los numera. Descripción corta, idealmente de una línea (hasta ~50 caracteres), con mayúscula inicial, sin la cantidad dentro del texto ("Limpieza de condensador y regulación de puertas"). La cantidad va en `cantidad`. Si una descripción es más larga, el script la parte en renglones, pero nunca meter un párrafo.
6. **Producto con componentes** (kit, unidad armada, equipo con accesorios): es UN SOLO ítem con sus componentes en `incluye` (sin precio ni cantidad propios). El script lo dibuja en negrita con "Incluye:" y viñetas, como un bloque que nunca se corta entre hojas. Si una parte la aporta el cliente (ej. el compresor), el trabajo va como componente ("Armado y conexión con el compresor del cliente") y se aclara en Observaciones (regla 11).
7. **Modelo**: vacío por defecto. Solo se completa si el párrafo dice un tipo, potencia o modelo específico (vitrina, cámara frigorífica, exhibidora, "1 HP", un modelo concreto). No poner "Heladera" genérico.
8. **Precio global**: si el precio es uno solo por todo el trabajo y hay varios ítems distintos, NO agregar un ítem de "mano de obra total": poner el importe en `precio_global` y dejar `precio_unitario` vacío (`null`) en todos los ítems; el script lo pone en el Subtotal del primer ítem. Si es un solo producto con componentes, el precio va en `precio_unitario` con su `cantidad`. Si el párrafo da precios por ítem (o "cada uno"), van en `precio_unitario`.
9. **IVA**: `0`, `21` o `10.5`, según el párrafo o la respuesta del Paso 0.
10. **Condición de pago**: SIEMPRE "100% contado".
11. **Observaciones**: vacías (`""`), salvo que Agustín diga qué poner o que el cliente aporte una parte del equipo o del trabajo: eso SIEMPRE se aclara (ej. "El compresor lo provee el cliente."). Si están vacías, el recuadro no aparece.

## Varias hojas (lo hace el script solo)

- Si todo entra en una hoja A4, sale una sola hoja y la tabla se completa con renglones vacíos hasta 7.
- Si no entra, corta ENTRE ítems (un producto nunca queda separado de sus componentes), escribe "Continúa en la hoja siguiente" y la hoja siguiente lleva un encabezado de continuación (logo chico, N° de presupuesto, "hoja 2 de 3", fecha) y repite el título de la tabla.
- Subtotal, IVA, TOTAL, condiciones y observaciones van siempre juntos al final y nunca quedan solos en una hoja: si no entran, se pasa con ellos al menos el último ítem.
- Todas las hojas llevan al pie "REFRIVAN · Documento no válido como factura." y "Hoja X de N".
- El script mide cómo imprime LibreOffice y verifica que el PDF tenga la cantidad de hojas calculada; si no, reajusta solo.

## Procedimiento

1. Paso 0: chequear cliente, ítems, precio e IVA; preguntar lo que falte.
2. Consultar el dólar BNA del día y el número siguiente.
3. Escribir el JSON (ej. `datos.json`) y correr:
   `python scripts/generar_presupuesto.py datos.json <carpeta de salida> --png`
   Deja `Presupuesto N<numero>.xlsx`, `.pdf` y un PNG por hoja.
4. MIRAR cada PNG antes de entregar. Auditar: logo y franja, número, fecha, cotización, un solo N° por producto, descripciones sin cortar feo, subtotal/IVA/total (miles con punto), observaciones y, si hay varias hojas, que el corte tenga sentido. Si algo se ve mal, corregir el JSON (acortar descripciones, agrupar componentes) y volver a correr.
5. Guardar los DOS archivos (xlsx y pdf) en la carpeta Presupuestos de la PC.
6. Resumir en pocas líneas qué se cargó, qué quedó vacío y cuántas hojas salieron.
7. Actualizar el contador en memoria.

Hablarle en castellano rioplatense, corto. Nada de tarjetas de opción múltiple: preguntar en texto normal, una cosa por vez.
