# Sistema de Punto de Venta (POS)

App de escritorio en Python (tkinter) para gestionar ventas y stock de un
comercio chico, con lector de código de barras, impresión de tickets en
impresora térmica USB, exportación a Excel y alertas por email cuando el
stock de un producto es bajo.

## 1. Instalación (Windows)

1. Instalá Python 3.10 o superior desde https://python.org (al instalar,
   marcá la casilla "Add Python to PATH").
2. Descomprimí esta carpeta en tu PC, por ejemplo en `C:\POS`.
3. Abrí una terminal (CMD o PowerShell) en esa carpeta y ejecutá:

   ```
   pip install -r requirements.txt
   ```

4. Para iniciar la app:

   ```
   python app.py
   ```

   Se va a crear automáticamente en la misma carpeta:
   - `stock.db` (base de datos de productos)
   - `config.json` (tu configuración)
   - `ventas.xlsx` (se crea la primera vez que guardás una venta)
   - `stock.xlsx` / `stock.txt` (se crean cuando generás el reporte de stock)

## 2. Lector de código de barras

Un lector de código de barras USB funciona como un teclado: al escanear,
"tipea" el código y aprieta Enter automáticamente. No necesita ningún
driver especial — solo tené el cursor parado en el campo "Escanear código
de barras" (en Nueva Venta) o "Código de barras" (en Administrar Stock).

## 3. Impresora térmica (USB)

1. Conectá la impresora e instalala normalmente en Windows (que aparezca
   en "Configuración > Impresoras y escáneres").
2. Si es tu impresora predeterminada, no necesitás configurar nada más.
3. Si no lo es, andá a **Configuración** dentro de la app y escribí el
   nombre EXACTO de la impresora tal como figura en Windows, en el campo
   "Nombre exacto de la impresora".
4. La app envía comandos ESC/POS crudos (estándar de casi todas las
   impresoras térmicas de tickets). Si tu impresora no imprime bien
   (caracteres raros, no corta el papel), contame el modelo y ajustamos
   los comandos.

**Importante:** esto requiere el paquete `pywin32` (ya está en
`requirements.txt`). Si no lo instalás, la app funciona igual pero el
botón "Imprimir ticket" te va a avisar que falta instalarlo.

## 4. Alertas de stock bajo por email

Como no ibas a usar SMS pago (Twilio), la app manda un **email** cuando
un producto queda con pocas unidades (por defecto, 3 o menos — lo podés
cambiar en Configuración).

Para que funcione, necesitás una cuenta de Gmail con "contraseña de
aplicación" (no tu contraseña normal, Google no lo permite por seguridad):

1. Activá la verificación en 2 pasos en tu cuenta de Google
   (myaccount.google.com/security).
2. Andá a https://myaccount.google.com/apppasswords y generá una
   contraseña de aplicación (elegí "Otra" y ponele un nombre, ej: "POS").
3. Copiá esa contraseña (16 caracteres) y pegala en Configuración, en el
   campo "Contraseña de aplicación de Gmail".
4. Completá también "Cuenta Gmail que envía" (esa misma cuenta de Gmail)
   y "Email que recibe el aviso" (puede ser tu propio email, tu celular
   con Gmail, etc.).
5. Guardá la configuración. A partir de ahí, cada vez que finalices una
   venta y algún producto quede en el umbral configurado o por debajo,
   te va a llegar un email automático.

Esta configuración queda guardada en `config.json` y se reutiliza cada
vez que abrís la app (no hay que cargarla de nuevo).

## 5. Cómo se usa

### Lobby
Tres botones grandes: **Nueva Venta**, **Administrar Stock** y
**Configuración**.

### Administrar Stock
- Escaneá o escribí un código de barras y apretá Enter: si el producto
  no existe, te deja completar Nombre, Precio y Stock inicial y guardarlo.
  Si ya existe, te trae sus datos para editarlos.
- La lista de abajo muestra todos los productos (nombre, precio, stock).
- Podés eliminar un producto seleccionándolo y usando el botón
  correspondiente.
- "Generar reporte de stock" crea `stock.xlsx` (todo el stock) y
  `stock.txt` (separado en "Disponibles" y "Sin stock").

### Nueva Venta
- Escaneá productos: se agregan a la lista automáticamente con nombre y
  precio (tomados de Administrar Stock). Si escaneás el mismo producto
  de nuevo, suma +1 a la cantidad.
- Podés editar la cantidad de un producto haciendo **doble clic** sobre
  el número de cantidad en la tabla.
- El campo "Margen de ganancia (%)" recalcula el TOTAL automáticamente:
  `TOTAL = Subtotal × (1 + margen / 100)`.
- Botones:
  - **Imprimir ticket**: manda el ticket a la impresora térmica.
  - **Guardar en Excel**: agrega la venta a `ventas.xlsx` (columnas: Día,
    Fecha, Hora, Producto, Precio Unitario, Cantidad, Subtotal).
  - **Finalizar compra**: descuenta el stock vendido, dispara los avisos
    por email si corresponde, y vuelve al lobby.
  - **Eliminar producto**: seleccioná un producto de la lista; si tiene
    más de 1 unidad cargada, te pregunta cuántas unidades sacar.

### Configuración
- Tema claro / oscuro.
- Tamaño de letra.
- Umbral de stock bajo.
- Datos de email para las alertas.
- Nombre de la impresora (opcional).

## 6. Notas sobre supuestos que tomé

- El **precio** que cargás en Administrar Stock es el que se usa como
  "precio unitario" en la venta. El **margen de ganancia** se aplica
  sobre el total de la compra para mostrar el precio final con ganancia
  (tal como lo describiste). Si en realidad querés que el margen se
  aplique producto por producto (no sobre el total), decímelo y lo
  ajusto — es un cambio chico.
- El Excel de ventas guarda una fila por producto vendido (no una fila
  por ticket completo), para que puedas después filtrar/sumar fácil en
  Excel.
- Si necesitás que funcione en Mac/Linux más adelante, la impresión
  térmica habría que cambiarla (pywin32 es solo Windows) — avisame.
