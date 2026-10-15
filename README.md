# Musculosa Los Guerreros · Playoff 2026

Página de pedidos (relevamiento) + backend en Google Apps Script que guarda todo en la planilla del club.

```
index.html            la página (todo en un archivo)
assets/front.png      frente de la musculosa
assets/back.png       espalda
assets/hinchada.jpg   fondo del inicio
assets/scrum.jpg      fondo de la pantalla final
apps-script/Code.gs   backend
```

## 1. Backend (una sola vez)

1. Abrí la planilla de Google Sheets del club.
2. **Extensiones > Apps Script**. Borrá lo que haya y pegá todo `apps-script/Code.gs`. Guardá.
3. (Opcional) Elegí `testSale` en el menú de funciones y tocá **Ejecutar**: en el registro tiene que decir `testSale OK`.
4. **Implementar > Nueva implementación** > tipo **Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
5. Autorizá los permisos (planilla y Drive). Copiá la URL que termina en **/exec**.
6. En `index.html`, arriba de todo en el `<script>`, pegala: `const API_URL = 'https://script.google.com/macros/s/.../exec';`

Con el primer pedido se crean solas las hojas **Pedidos** y **Resumen** (cantidades por talle y por división, y lo cobrado) y la carpeta de Drive **Comprobantes Musculosa 2026** (privada).

> Si cambiás `Code.gs` después: **Implementar > Administrar implementaciones > editar > Versión: nueva**. Así la URL no cambia.

## 2. Cómo se usa la planilla

- Cada pedido entra como **Pendiente de pago**. Cuando suben el comprobante pasa a **Comprobante recibido** con el link al archivo.
- Cuando el club verifica la transferencia, cambia el Estado a **Pagado** a mano (hay un desplegable). También existe **Cancelado**.
- **Sale** (terminó): valió para pedidos del 15/10/2026 con comprobante antes de las 23:59. Ahora la página muestra precio fijo de $38.000; el backend ya no aplica descuento.
- La página ya no pide apodo: en la columna *Apodo* se guarda el nombre y apellido.
- Precio, descuento, fecha del sale y alias se cambian en las constantes de arriba de `Code.gs`.
- La columna oculta **Clave** protege el link de cada pedido; no la borres.
- Para bajar un Excel: **Archivo > Descargar > Microsoft Excel (.xlsx)**.

## 3. Publicar la página gratis

Cualquiera de estas (subís la carpeta con `index.html` y `assets/`; `apps-script/` y este README no hacen falta):

- **Netlify Drop**: entrá a https://app.netlify.com/drop y arrastrá la carpeta. Te da un link al toque.
- **GitHub Pages**: subí los archivos a un repo > Settings > Pages > Deploy from branch `main`.
- **Cloudflare Pages**: Workers & Pages > Create > Pages > Upload assets > subí la carpeta.

Para probar en la compu: `python -m http.server` dentro de la carpeta y abrí http://localhost:8000.
Si `API_URL` está vacío, o abrís la página con `?demo=1` al final de la URL, funciona en **modo demo** (simula los pedidos en el navegador, no guarda nada en la planilla).
