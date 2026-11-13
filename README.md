# Pólizas flota vehicular

Aplicación web para prorratear la factura semestral de la póliza de flota de La Segunda entre las patentes y ver el monto total por centro de costo (CC).

Cada semestre se suben la póliza (PDF), la factura (PDF) y, si cambiaron, la base de CC y la base de zonas (Excel). La aplicación lee los archivos en el navegador, controla que cuadren y guarda todo en Supabase. Las patentes que ya tenían CC lo conservan en el período siguiente.

## Stack

- Vite + JavaScript, sin framework. Se despliega en Vercel como sitio estático.
- Supabase: Postgres con RLS. Se ingresa solo con el correo, que tiene que estar en la tabla `seg_usuarios` (sin contraseña).
- `pdfjs-dist` para leer los PDF y SheetJS (`xlsx`) para leer y exportar Excel. SheetJS se instala desde su CDN oficial, porque la versión publicada en npm está desactualizada.

## Puesta en marcha

### 1. Supabase

1. En el proyecto, abrí **SQL Editor**, pegá el contenido completo de `supabase/migrations/20261112120000_seguros_cc.sql` y ejecutalo. Crea las tablas, las vistas y las funciones. Después ejecutá `supabase/migrations/20261112150000_prorrateo_solo_premio.sql`, que fija el prorrateo por premio, `supabase/migrations/20261113120000_cc_solo_habilitados.sql`, que impide guardar asignaciones con CC inhabilitados, y `supabase/migrations/20261114120000_factura_pdf.sql`, que guarda el PDF de cada factura. Todos se pueden ejecutar más de una vez sin perder datos. Si usás la CLI, alcanza con `supabase db push`.
2. Ejecutá `supabase/usuarios.sql` para cargar los correos habilitados.
   - Para agregar a alguien, sumá una línea en ese archivo y volvé a ejecutarlo.
   - Para dar de baja a alguien: `update public.seg_usuarios set activo = false where email = 'correo@fabricasrl.com.ar';`
3. Copiá la **Publishable key** desde **Project Settings → API Keys**. Si el proyecto todavía usa las claves anteriores, la `anon` key funciona igual.

### 2. GitHub y Vercel

1. Subí esta carpeta a un repositorio nuevo:
   ```bash
   git init && git add . && git commit -m "Primera versión"
   git branch -M main
   git remote add origin https://github.com/<usuario>/seguros-cc.git
   git push -u origin main
   ```
2. En Vercel: **Add New → Project**, importá el repositorio. Vercel detecta Vite solo.
3. En **Environment Variables** cargá:
   - `VITE_SUPABASE_URL` = `https://oixzkvkybqexvrzwzwis.supabase.co`
   - `VITE_SUPABASE_PUBLISHABLE_KEY` = la clave del paso 1.5
4. Hacé clic en **Deploy**. Si cambiás las variables después, hay que volver a desplegar, porque se incrustan en el build.

### Desarrollo local

```bash
cp .env.example .env.local   # y completá la clave
npm install
npm run dev
```

Requiere Node 22.12 o superior.

## Uso

1. Abrí la URL de la aplicación e ingresá con tu correo de la empresa. La app lo recuerda en ese navegador hasta que tocás «Salir».
2. **Cargar nuevo período** → arrastrá los archivos juntos. La aplicación reconoce cada uno:
   - **Póliza** (PDF): de cada certificado toma patente, marca completa, año, valor del vehículo y premio.
   - **Factura** (PDF): toma neto, IVA, percepción, otros impuestos y total.
   - **Base de CC** (Excel con columnas `CC`, `Descripción` y `Habilitado` en la columna C). Solo los CC con `VERDADERO` en «Habilitado» se pueden asignar.
   - **Zonas** (Excel con columnas `Patente`, `Zona`), opcional.
3. Revisá los controles (cantidad de certificados, total de factura contra premio final de la póliza, patentes sin zona) y guardá.
4. El **Dashboard** es la primera pestaña y resume el período visible con los datos guardados:
   - Indicadores: costo total de la factura, vehículos asegurados, centros de costo con vehículos, porcentaje de patentes con CC habilitado, patentes sin asignación y días de vigencia que quedan.
   - Costo y participación por centro de costo (barras y anillo). Aparecen todos los CC del período, ordenados por monto, con el nombre completo y el importe al centavo, para cargar el prorrateo en Bejerman. La leyenda del anillo cierra con el total. Si tocás una barra, se abre ese CC en **Totales por CC**.
   - Estado de patentes, últimos períodos con la variación del total facturado contra el período anterior (el ojo abre ese período), y alertas: patentes sin CC o con CC inhabilitado, diferencia entre factura y prorrateo, patentes nuevas o sin zona, y la última carga de la base de CC. Las alertas con flecha llevan a las patentes correspondientes.
5. El ícono de la lupa con el PDF, arriba a la derecha de la tarjeta de la factura (y «Ver factura» en el Dashboard), abre el PDF de la factura en un visor con zoom, para descargar o abrir en otra pestaña. El PDF se guarda al cargar el período. Si el período se cargó antes, el ícono muestra un «+» y permite subirlo; la app controla que sea la misma factura antes de guardarlo.
6. En **Totales por CC** está el monto por centro de costo. Si tocás un CC, se ven sus patentes.
7. En **Patentes y asignación** se asigna el CC: fila por fila, o en bloque filtrando por zona y seleccionando todas. El botón con la flecha abre un menú como el autofiltro de Excel: un buscador por código o nombre (sin importar mayúsculas ni tildes), la lista de CC habilitados con un casillero cada uno, y **Aceptar** / **Cancelar**. Con el teclado: ↓ o F4 abre el menú, ↑ ↓ recorren la lista, Enter acepta la opción marcada y Esc cancela. Si una patente tiene un CC que quedó inhabilitado, se marca en naranja y el menú pide elegir otro. También se puede importar un Excel con columnas `Patente` y `CC`; las filas con CC inhabilitados se omiten y se informan.
8. **Exportar a Excel** genera tres hojas: totales por CC, detalle por patente y datos de la factura.

El prorrateo se hace por **valor del vehículo** por defecto. Se puede cambiar a **premio del certificado** con el selector. Cada concepto se reparte al centavo con el método del mayor resto, así cada columna suma exactamente lo facturado.

## Íconos

Los íconos están en `public/`: `favicon.ico` y `favicon.svg` (escudo simplificado, legible en la pestaña), `icon.svg` (vehículo con escudo), `apple-touch-icon.png`, `icon-192.png`, `icon-512.png` e `icon-maskable-512.png` para Android. `manifest.webmanifest` permite instalar la app en el celular o el escritorio. Los colores son los del logo de la empresa: azul `#0A4088` y rojo `#BA1212`.

## Colores

La paleta está definida una sola vez en las variables de `:root` al comienzo de `src/style.css`, con una versión para modo claro y otra para modo oscuro, y la usan todas las pantallas en escritorio y celular. El verde azulado oscuro `#1E4952` se usa en títulos y encabezados de tarjetas, y el verde azulado `#1D7A70` en botones y enlaces. Los gráficos usan la serie `--s1` a `--s12`: `#2A9D8F`, `#264653`, `#8AB17D`, `#E9A23B`, `#E76F51`, `#9DB4C0`, `#5E8C6A`, `#B07AA1`, `#4F7CAC`, `#8E6C4A`, `#7A5C99` y `#D9A6A0`. Si hay más de doce CC, los colores se repiten en tonos más claros y más oscuros. Verde, naranja y rojo indican estado: correcto, pendiente y error.

## Modelo de datos

| Tabla / vista | Contenido |
| --- | --- |
| `seg_usuarios` | Correos habilitados para ingresar (con nombre y estado activo). |
| `seg_centros_costo` | Base de CC vigente (se reemplaza completa al subir una nueva). |
| `seg_zonas` | Zona por patente (se reemplaza completa al subir una nueva). |
| `seg_asignaciones` | Último CC conocido por patente; precarga los períodos nuevos. |
| `seg_bases` | Archivo y fecha de la última carga de cada base. |
| `seg_periodos` | Un registro por factura: póliza, vigencia e importes. |
| `seg_facturas_pdf` | PDF de la factura de cada período. Solo se lee y se guarda con `seg_factura_pdf` y `seg_guardar_factura_pdf`. |
| `seg_periodo_vehiculos` | Vehículos del período con su CC. |
| `seg_v_prorrateo` | Vista con el prorrateo por patente (neto, IVA, percepción, otros, total). |
| `seg_v_totales_cc` | Vista con los totales por período y CC, incluidas las zonas. |

Las vistas sirven también para consultar desde otros sistemas, por ejemplo:

```sql
select cc, descripcion, total
from seg_v_totales_cc
where periodo_id = '0027-00233598'
order by total desc;
```

Todas las escrituras pasan por funciones transaccionales (`seg_guardar_periodo`, `seg_asignar_cc`, `seg_cambiar_metodo`, `seg_reemplazar_centros_costo`, `seg_reemplazar_zonas`).

## Seguridad

Se ingresa solo con el correo, sin contraseña. La app envía ese correo a Supabase en cada consulta (encabezado `x-usuario`), y la base solo muestra datos y acepta cambios si está en `seg_usuarios` y activo. Además, cada período y cada asignación registran qué correo los hizo.

Esto controla quién usa la app, pero no es una autenticación fuerte: alguien que conozca un correo habilitado y la URL podría ingresar. Para acotar el riesgo:

- Las tablas son de solo lectura para la API. No se pueden borrar ni editar filas directamente; solo a través de las funciones de la aplicación.
- Conviene no publicar la URL fuera de la empresa.
- Si más adelante hace falta más seguridad sin contraseñas, Supabase permite ingresar con un enlace que llega al correo (*magic link*), y Vercel ofrece **Deployment Protection** en **Settings → Deployment Protection**.

## Si cambia el formato de la póliza

La lectura busca los rótulos «Valor del vehículo», «Marca completa», «Patente», «Año», «Certificado N°» y «PREMIO TOTAL» en `src/parsers.js`. Si La Segunda cambia el diseño del PDF, los controles de carga lo marcan antes de guardar, y el ajuste se hace en ese archivo.
