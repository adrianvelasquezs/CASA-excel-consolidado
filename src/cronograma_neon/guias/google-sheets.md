# Versión Google Sheets

Mismo cronograma consolidado, pero con permisos **por cuenta institucional** en lugar de contraseñas. Es la única de las dos versiones que cumple de forma real que lo que registra un profesor no lo pueda modificar otro.

Pestaña **Google Sheets** de la app. **Qué se carga y qué se guarda** (siempre donde usted elija):

| | Botón | Archivo |
|---|---|---|
| Entrada | *Elegir archivos… / Elegir carpeta…* (arriba) | Plantillas .xlsx de los programas (o las incluidas en la app) |
| Salida | *Generar plantilla para Google Sheets…* | Plantilla que se importa a Google Sheets (sin contraseñas) |
| Salida | *Guardar Codigo.gs… / Copiar Codigo.gs* | Apps Script que aplica los permisos, con la configuración de la sección 2 ya aplicada |
| Salida | *Guardar appsscript.json… / Copiar appsscript.json* | Manifiesto del script (zona horaria America/Bogota) |

## Cómo funciona

**Roles.** Se definen en la hoja **Usuarios**, una fila por correo:

| Rol | Puede editar |
|---|---|
| `ADMIN` | Todo. Es el único que usa el menú **CASA** |
| `CASA` | Columnas verdes de todas las hojas de programa |
| `GESTOR` | Columnas azules de su(s) programa(s), incluidas las filas de sus profesores |
| `PROFESOR` | Columnas azules de su(s) programa(s). Cuando registra una fila, solo esa persona puede volver a editarla (además de los gestores y ADMIN) |

En la columna Programa(s), varios programas van separados por coma (`AFIN, MF`).

**Tres capas de control:**
1. **Protecciones de Google Sheets por correo.** Encabezados, columnas del sistema y las hojas Usuarios, Consolidado e Instrucciones quedan solo para ADMIN. La zona azul de cada hoja queda para el programa y la zona verde para CASA. Google bloquea en la propia interfaz cualquier edición fuera de esto.
2. **Propiedad de filas.** Cuando un profesor escribe en una fila vacía, el script anota su correo y la fecha en **Registrado por / Fecha de registro** y protege la fila para esa persona. Si la persona vacía todo el contenido azul de la fila, la fila se libera.
3. **Red de seguridad.** En Google Sheets, quien puede editar un rango protegido también puede quitarle la protección. Por eso el script revisa cada edición contra los roles. Si alguien logra editar algo que no le corresponde, el script:
   - revierte la edición (si fue una sola celda),
   - la anota en la hoja oculta **_Auditoria**,
   - avisa por correo a los ADMIN,
   - restaura las protecciones de esa hoja.

   Además, cada noche se restauran todas las protecciones (a las 2 a. m. por defecto; la hora se cambia en la app).

## Particularidades de Uniandes (Google Workspace + correo en Microsoft)

- **Entrar con la cuenta institucional.** Cada persona debe abrir el archivo con su cuenta de Google `@uniandes.edu.co`, que normalmente se ingresa con las mismas credenciales de Microsoft. Si el navegador tiene abierta una cuenta personal de Gmail, Google mostrará "Solicitar acceso". En ese caso hay que cambiar de cuenta (foto de perfil, arriba a la derecha).
- **Solo correos `@uniandes.edu.co`.** Google solo le dice al script quién editó si la cuenta pertenece al mismo dominio que la cuenta que instaló el script. Una cuenta externa (Gmail personal, otra universidad) no se puede identificar: sus ediciones se revierten. Al aplicar permisos, el script avisa si hay correos externos en Usuarios.
- **Las notificaciones llegan a Outlook, si TI lo tiene configurado así.** Los buzones viven en Microsoft 365. Las notificaciones de Google (archivo compartido, menciones, alertas del script) solo llegan a Outlook si en la consola de Google Workspace el servicio Gmail está desactivado para esos usuarios, o si el enrutamiento está configurado hacia Microsoft. De lo contrario quedan en un buzón de Gmail que nadie revisa. Pruébelo con una alerta antes de depender de ella, y comparta el enlace del archivo también por Teams.
- **Grupos.** Las listas de distribución y los grupos de Microsoft 365 no existen en Google. Registre en Usuarios a cada persona con su correo individual.
- **Políticas de TI.** Si la consola de Workspace restringe Apps Script o las apps no verificadas, puede que la autorización del paso 4 quede bloqueada. En ese caso, TI debe permitir el proyecto de Apps Script (es interno del dominio).

## Puesta en marcha (una sola vez)

Hágalo con una cuenta **@uniandes.edu.co de CASA**, idealmente la que será dueña del archivo. Esa cuenta queda como ADMIN automático y es la que ejecuta los disparadores.

1. **Generar la plantilla.** En la app, pestaña *Google Sheets*: revise el dominio y las opciones del Apps Script, clic en **Generar plantilla para Google Sheets…** y elija dónde guardarla.
2. **Importar a Google Sheets.** En Google Drive: *Nuevo → Subir archivo* y luego *Abrir con → Hojas de cálculo de Google*. O bien en Sheets: *Archivo → Importar → Subir*, eligiendo "Reemplazar hoja de cálculo". Trabaje sobre el archivo de Google Sheets, no sobre el .xlsx.
3. **Llenar la hoja Usuarios.** Incluya al menos a los ADMIN y al equipo CASA. Si el archivo está en una **unidad compartida** (sin dueño), agréguese usted como `ADMIN` antes del paso 5.
4. **Instalar el script.** En el archivo: *Extensiones → Apps Script*.
   - En la app, clic en **Copiar Codigo.gs** y pegue el texto reemplazando todo el contenido de `Código.gs`.
   - En ⚙ *Configuración del proyecto*, marque "Mostrar el archivo de manifiesto appsscript.json". En la app, clic en **Copiar appsscript.json** y péguelo en ese archivo.
   - Guarde y recargue la hoja de cálculo. Aparece el menú **CASA**.
   - (Opcional) Con [clasp](https://github.com/google/clasp): guarde los dos archivos con *Guardar Codigo.gs…* y *Guardar appsscript.json…* en la carpeta de `clasp clone <scriptId>` y ejecute `clasp push`.
5. **CASA → 1. Configurar libro.** Autorice los permisos que pide Google. Esto escribe la fórmula del Consolidado, crea las protecciones e instala los disparadores. Revise el resumen: avisa de roles inválidos, programas inexistentes y correos externos.
6. **CASA → 2. Compartir con los correos de Usuarios.** Da acceso de edición a todos los de la hoja. Las protecciones ya están puestas, así que cada quien solo podrá modificar lo suyo.
7. **Restringir quién comparte.** Botón *Compartir* → ⚙ → desmarque "Los editores pueden cambiar los permisos y compartir".

## Operación diaria

- **Alta, baja o cambio de rol:** edite la hoja Usuarios y ejecute **CASA → Aplicar permisos ahora**. Para un usuario nuevo, ejecute también *Compartir*.
- **Revisar incidentes:** **CASA → Ver auditoría**.
- **Recuperar contenido** (por ejemplo, un pegado de varias celdas que no se pudo revertir): *Archivo → Historial de versiones*, o clic derecho en una celda → *Mostrar historial de ediciones*.
- **Más filas:** agregue filas al final de la hoja del programa, copie la fórmula de la columna Programa hacia abajo y ejecute *Aplicar permisos ahora*.
- **Filtrar:** use *Datos → Vistas de filtro*. El filtro normal está bloqueado porque cambiaría la vista de todos.

## Prueba piloto antes de abrirlo a todos

El script se probó con una **simulación** de Google Sheets: 37 casos, entre ellos permisos por rol, propiedad y liberación de filas, reversión de ediciones indebidas, restauración de protecciones, re-sincronización sin duplicados y menú solo para ADMIN. **No se ha ejecutado en Google Sheets real.** Esa simulación supone que, cuando una celda está cubierta por varias protecciones, Google exige ser editor de todas ellas. Verifíquelo con 3 o 4 cuentas reales:

| Prueba | Resultado esperado |
|---|---|
| Profesor A escribe en una fila vacía de su programa | Se llena "Registrado por" con su correo |
| Profesor B (mismo programa) intenta editar esa fila | Google lo bloquea |
| Gestor del programa edita esa fila | Puede |
| CASA edita una columna verde de esa fila | Puede |
| CASA intenta editar una columna azul | Google lo bloquea |
| Profesor de otro programa intenta editar la hoja | Google lo bloquea |
| Cualquiera que no sea ADMIN intenta editar encabezados, Usuarios o Consolidado | Google lo bloquea |
| ADMIN ejecuta *Ver auditoría* tras los intentos | No hay incidentes; los bloqueos de Google no llegan al script |

Si CASA no puede editar las columnas verdes de una fila registrada, la suposición anterior no se cumple. Avíseme para cambiar la protección de las filas a solo las columnas azules (genera más protecciones).

## Limitaciones

- **Todos ven todo.** Cualquiera con acceso al archivo puede leer los datos de todos los programas, la hoja Usuarios incluida. Los permisos solo controlan quién modifica. Si un programa no debe ver a otro, hace falta un archivo por programa y un consolidado aparte con `IMPORTRANGE`.
- **Pegados de varias celdas.** Si alguien quita una protección y pega varias celdas, el script no puede restaurar los valores anteriores: lo registra, avisa y restaura las protecciones. La recuperación se hace con el historial de versiones.
- **Margen de segundos.** Entre que un profesor escribe en una fila libre y el script la protege pasan unos segundos.
- **Un solo instalador.** Los disparadores corren con la cuenta que ejecutó *Configurar libro*. Si esa cuenta se desactiva, otro ADMIN debe volver a ejecutarlo; primero hay que borrar los disparadores anteriores en Apps Script → *Activadores*.
