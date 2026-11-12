# Seguros automotor por centro de costo

Aplicación web para prorratear la factura semestral de la póliza de flota de La Segunda entre las patentes y ver el monto total por centro de costo (CC).

Cada semestre se suben la póliza (PDF), la factura (PDF) y, si cambiaron, la base de CC y la base de zonas (Excel). La aplicación lee los archivos en el navegador, controla que cuadren y guarda todo en Supabase. Las patentes que ya tenían CC lo conservan en el período siguiente.

## Stack

- Vite + JavaScript, sin framework. Se despliega en Vercel como sitio estático.
- Supabase: Postgres con RLS y Auth con email y contraseña.
- `pdfjs-dist` para leer los PDF y SheetJS (`xlsx`) para leer y exportar Excel. SheetJS se instala desde su CDN oficial, porque la versión publicada en npm está desactualizada.

## Puesta en marcha

### 1. Supabase

1. En el proyecto, abrí **SQL Editor** y ejecutá el contenido de `supabase/migrations/20261112120000_seguros_cc.sql`. Si usás la CLI, alcanza con `supabase db push`.
2. Creá los usuarios en **Authentication → Users → Add user**, con la opción *Auto Confirm User*.
3. Habilitalos en la aplicación. Solo los emails cargados en esta tabla pueden ver o modificar datos:
   ```sql
   insert into public.seg_usuarios (email, nombre) values
     ('nombre@empresa.com', 'Nombre Apellido');
   ```
4. Desactivá el registro público en **Authentication → Sign In / Providers → Allow new users to sign up**.
5. Copiá la **Publishable key** desde **Project Settings → API Keys**. Si el proyecto todavía usa las claves anteriores, la `anon` key funciona igual.

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
5. En Supabase, en **Authentication → URL Configuration**, poné la URL de Vercel como *Site URL*.

### Desarrollo local

```bash
cp .env.example .env.local   # y completá la clave
npm install
npm run dev
```

Requiere Node 22.12 o superior.

## Uso

1. Ingresá con tu usuario.
2. **Cargar nuevo período** → arrastrá los archivos juntos. La aplicación reconoce cada uno:
   - **Póliza** (PDF): de cada certificado toma patente, marca completa, año, valor del vehículo y premio.
   - **Factura** (PDF): toma neto, IVA, percepción, otros impuestos y total.
   - **Base de CC** (Excel con columnas `CC`, `Descripción`, `Habilitado`).
   - **Zonas** (Excel con columnas `Patente`, `Zona`), opcional.
3. Revisá los controles (cantidad de certificados, total de factura contra premio final de la póliza, patentes sin zona) y guardá.
4. En **Totales por CC** está el monto por centro de costo. Si tocás un CC, se ven sus patentes.
5. En **Patentes y asignación** se asigna el CC: fila por fila, o en bloque filtrando por zona y seleccionando todas. También se puede importar un Excel con columnas `Patente` y `CC`.
6. **Exportar a Excel** genera tres hojas: totales por CC, detalle por patente y datos de la factura.

El prorrateo se hace por **valor del vehículo** por defecto. Se puede cambiar a **premio del certificado** con el selector. Cada concepto se reparte al centavo con el método del mayor resto, así cada columna suma exactamente lo facturado.

## Modelo de datos

| Tabla / vista | Contenido |
| --- | --- |
| `seg_usuarios` | Emails habilitados. |
| `seg_centros_costo` | Base de CC vigente (se reemplaza completa al subir una nueva). |
| `seg_zonas` | Zona por patente (se reemplaza completa al subir una nueva). |
| `seg_asignaciones` | Último CC conocido por patente; precarga los períodos nuevos. |
| `seg_bases` | Archivo y fecha de la última carga de cada base. |
| `seg_periodos` | Un registro por factura: póliza, vigencia e importes. |
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

Todas las escrituras pasan por funciones transaccionales (`seg_guardar_periodo`, `seg_asignar_cc`, `seg_cambiar_metodo`, `seg_reemplazar_centros_costo`, `seg_reemplazar_zonas`). Estas funciones verifican que el usuario esté en `seg_usuarios`.

## Si cambia el formato de la póliza

La lectura busca los rótulos «Valor del vehículo», «Marca completa», «Patente», «Año», «Certificado N°» y «PREMIO TOTAL» en `src/parsers.js`. Si La Segunda cambia el diseño del PDF, los controles de carga lo marcan antes de guardar, y el ajuste se hace en ese archivo.
