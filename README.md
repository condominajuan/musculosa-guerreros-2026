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
3. Elegí `configurarPlanilla` en el menú de funciones y tocá **Ejecutar** (una sola vez; autorizá los permisos). Agrega la casilla *Pago confirmado* en Pedidos y arma las hojas **Todos** y **Dashboard**.
4. **Implementar > Nueva implementación** > tipo **Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
5. Autorizá los permisos (planilla y Drive). Copiá la URL que termina en **/exec**.
6. En `index.html`, arriba de todo en el `<script>`, pegala: `const API_URL = 'https://script.google.com/macros/s/.../exec';`

La carpeta de Drive **Comprobantes Musculosa 2026** (privada) se crea con el primer comprobante.

> Si cambiás `Code.gs` después: **Implementar > Administrar implementaciones > editar > Versión: nueva**. Así la URL no cambia.

## 2. Cómo se usa la planilla

- **Flujo web:** la persona completa sus datos (no se anota nada todavía), transfiere y sube el comprobante. Recién ahí el pedido se anota en **Pedidos Web**, con su ID (G-0001...), el estado **Comprobante recibido** y el link al archivo. Los pedidos web ya no pasan por *Pendiente de pago* (ese estado queda para el Forms y para destildar).
- Si alguien cierra la página para ir al home banking, al volver (en el mismo celular/navegador) sigue en el paso del pago con sus datos. No hay link de pedido para retomarlo desde otro dispositivo.
- Si el envío se corta y lo reintentan, no se duplica: cada reserva manda un código único (válido 6 horas).
- La columna vieja **Clave** ya no se usa: podés borrarla o dejarla vacía.
- Cuando el club verifica la transferencia, tilda la casilla **Pago confirmado** y el Estado pasa solo a *Pago confirmado* (destildar lo devuelve). El Estado también tiene *Cancelado* (desplegable).
- **Todos** junta los pedidos del Forms del sale (hoja *Respuestas de formulario 1*, la columna con el check = pagó) y los de la web, con filtros en la fila 1.
- **Dashboard**: cantidad de musculosas por división y talle (todas las reservas y solo pagos confirmados) y estado de pagos por división.
- El precio ($38.000) y el alias se cambian en las constantes de arriba de `Code.gs`.
- **Avisar por WhatsApp**: al tildar *Pago confirmado* aparece en la columna *Avisar* un link que abre WhatsApp con el mensaje armado (se edita en `mensaje_` de `Code.gs`). Los celulares se normalizan a formato 549 + área + número. Si el número no sirve dice *Revisar número*. Con la casilla *Aviso enviado* se marca a quién ya se le escribió.
