# GestionTIC

Aplicación web en **Google Apps Script** para registrar visitas de soporte TIC desde un QR.

## Flujo implementado

1. El usuario abre la web desde QR.
2. Ingresa **carnet o cédula** y busca.
3. Si existe en `Usuarios`, se autocompletan datos.
4. Si no existe, se crea el usuario con datos personales y clave.
5. Registra la visita en `Registros` con:
   - Fecha
   - Hora
   - Descripción (motivo)
   - Nombre y apellido
   - Referencia (carnet/cédula)
   - Observación
   - Firma (imagen embebida en celda con `=IMAGE(...)`, no solo link)
6. Se muestra histórico de gestiones por referencia.

## Hojas requeridas

La app crea/ajusta automáticamente:

- `Usuarios`
- `Registros`

## Reglas aplicadas

- Si coincide con formato de cédula `000-000000-0000A`, se clasifica como **maestro**.
- Si no coincide, se asume **alumno**.
- Compatibilidad con `tipo_usuario` heredado en hoja `Usuarios`:
  - `A`, `ALUMNO`, `ESTUDIANTE` => **alumno**
  - `M`, `MAESTRO`, `DOCENTE`, `PROFESOR` => **maestro**
- Fecha de nacimiento:
  - Maestro: derivada desde la cédula (bloque `ddmmaa`).
  - Alumno: ingresada manualmente.
- Control anti-registros basura: cada referencia usa una **clave (pin)** que se exige en visitas futuras.

## Carpeta de firmas en Drive

- En `Code.gs`, configura `CONFIG.SIGNATURE_FOLDER_ID`.
- Puedes poner **solo el ID** de la carpeta o **la URL completa** de Google Drive.
- En este proyecto ya está configurada la carpeta compartida:
  - `https://drive.google.com/drive/folders/10yZgZY3MfnCQAhh4RtTuusQ-RjDWUUu-`
- Si no se configura, la app usará (o creará) una carpeta llamada `Firmas_GestionTIC`.

## Archivos

- `Code.gs`: lógica backend, validaciones, Google Sheets y Drive.
- `index.html`: interfaz de búsqueda, alta, registro y firma táctil.
- `appsscript.json`: configuración del proyecto Apps Script.

## Despliegue rápido

1. Crear un proyecto en [script.google.com](https://script.google.com).
2. Copiar `Code.gs`, `index.html` y `appsscript.json`.
3. Vincular a una Google Sheet (o abrir desde una Sheet con Apps Script).
4. Implementar como **Aplicación web**:
   - Ejecutar como: tú.
   - Acceso: quienes tengan el enlace (o según necesidad).
5. Generar QR con la URL de la web app publicada.
