# Cronograma consolidado en Google Sheets

Un solo libro de Google Sheets con una hoja por programa, permisos **por cuenta institucional** y una hoja «Consolidado» que une todo. Lo que registra un profesor queda a su nombre y ningún otro profesor puede modificarlo.

## Configurar en la app

El libro se genera **desde cero** con lo que se define en la app; no hace falta ningún Excel de entrada.

| Pestaña | Qué se define |
|---|---|
| **Columnas y permisos** | Las columnas de cada hoja de programa, en orden. Para cada una: nombre, tipo (texto, número, fecha, SI/NO o lista de opciones), ancho, **quién puede editarla** (Profesores, Gestores, CASA; ADMIN siempre) y si **identifica el curso** |
| **Programas** | Las hojas del libro, una por programa |
| **Opciones** | Filas de registro por programa y opciones del Apps Script (dominio, hora de restauración, alertas, gestores) |

Para cambiar una columna: selecciónela en la tabla, ajuste el formulario y use *Guardar cambios de la columna*. Para crear una: llene el formulario y use *Agregar como columna nueva* (queda debajo de la seleccionada). *↑ Subir* y *↓ Bajar* cambian el orden. Todo se guarda solo en las preferencias de la app.

**Compartir la configuración con el equipo:** use *Exportar configuración…* (genera un `.json`) y la otra persona usa *Importar configuración…*. *Restablecer predeterminada* vuelve a los 12 programas y columnas originales del Cronograma 2027.

**Qué se guarda** (siempre donde usted elija):

| Botón | Archivo |
|---|---|
| *Generar plantilla para Google Sheets…* | Plantilla .xlsx que se importa a Google Sheets |
| *Guardar Codigo.gs… / Copiar Codigo.gs* | Apps Script que aplica los permisos, con las opciones ya aplicadas |
| *Guardar appsscript.json… / Copiar appsscript.json* | Manifiesto del script (zona horaria America/Bogota) |

## Columnas y colores

El color del encabezado indica quién edita:

| Color | Quién edita |
|---|---|
| Azul | Profesores y/o gestores del programa |
| Verde | Equipo CASA |
| Amarillo | El programa y el equipo CASA |
| Gris | Solo ADMIN (incluye «Programa», «Registrado por» y «Fecha de registro», que llena el sistema) |

La configuración predeterminada incluye, al final y solo para CASA: **¿Se entregó reporte?** (SI/NO), **Razones de porqué están pendientes reportes** y **Razones de porqué no se realizó medición**.

## Hoja «Permisos»

El libro trae una hoja **Permisos** (visible para todos, editable solo por ADMIN) con una fila por columna: Profesores / Gestores / CASA / Identifica el curso, con SI o NO. **El Apps Script toma los permisos de esta hoja**, no de los colores.

- **Cambiar permisos con el libro en uso:** edite SI/NO en esa hoja y luego *CASA → Aplicar permisos ahora*. No cambie los nombres de las columnas: una columna que no aparece en Permisos queda solo para ADMIN, y el script lo avisa.
- **Agregar, quitar o reordenar columnas:** genere un libro nuevo desde la app. El libro en uso no cambia de estructura.

## Cómo funciona

**Roles.** Se definen en la hoja **Usuarios**, una fila por correo:

| Rol | Puede editar |
|---|---|
| `ADMIN` | Todo. Es el único que usa el menú **CASA** |
| `CASA` | Las columnas marcadas para CASA en todas las hojas de programa |
| `GESTOR` | Las columnas marcadas para Gestores en su(s) programa(s), incluidas las filas de sus profesores |
| `PROFESOR` | Las columnas marcadas para Profesores en su(s) programa(s). Cuando registra una fila, solo esa persona puede volver a editarla (además de los gestores, CASA en sus columnas y ADMIN) |

En la columna Programa(s), varios programas van separados por coma (`AFIN, MF`).

**Tres capas de control:**
1. **Protecciones de Google Sheets por correo.** Encabezados, columnas del sistema y las hojas Usuarios, Consolidado, Permisos e Instrucciones quedan solo para ADMIN. Cada grupo de columnas queda para los roles marcados en la hoja Permisos. Google bloquea en la propia interfaz cualquier edición fuera de esto.
2. **Propiedad de filas.** Cuando un profesor escribe en una fila vacía, el script anota su correo y la fecha en **Registrado por / Fecha de registro** y protege la fila para esa persona. Si la persona vacía todo lo que escribió en las columnas del programa, la fila se libera.
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

1. **Generar la plantilla.** En la app, revise columnas, permisos, programas y opciones; luego clic en **Generar plantilla para Google Sheets…** y elija dónde guardarla.
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

El script se probó con una **simulación** de Google Sheets: 57 casos, entre ellos permisos por rol y por columna, cambios de permisos en vivo desde la hoja Permisos, propiedad y liberación de filas, reversión de ediciones indebidas, restauración de protecciones, re-sincronización sin duplicados y menú solo para ADMIN. **No se ha ejecutado en Google Sheets real.** Esa simulación supone que, cuando una celda está cubierta por varias protecciones, Google exige ser editor de todas ellas. Verifíquelo con 3 o 4 cuentas reales:

| Prueba | Resultado esperado |
|---|---|
| Profesor A escribe en una fila vacía de su programa | Se llena "Registrado por" con su correo |
| Profesor B (mismo programa) intenta editar esa fila | Google lo bloquea |
| Gestor del programa edita esa fila | Puede |
| CASA edita una columna de CASA (verde) de esa fila | Puede |
| CASA intenta editar una columna solo del programa (azul) | Google lo bloquea |
| Profesor de otro programa intenta editar la hoja | Google lo bloquea |
| Cualquiera que no sea ADMIN intenta editar encabezados, Usuarios o Consolidado | Google lo bloquea |
| ADMIN ejecuta *Ver auditoría* tras los intentos | No hay incidentes; los bloqueos de Google no llegan al script |

Si CASA no puede editar las columnas verdes de una fila registrada, la suposición anterior no se cumple. Hay que cambiar la protección de las filas a solo las columnas del programa (genera más protecciones).

## Limitaciones

- **Todos ven todo.** Cualquiera con acceso al archivo puede leer los datos de todos los programas, la hoja Usuarios incluida. Los permisos solo controlan quién modifica. Si un programa no debe ver a otro, hace falta un archivo por programa y un consolidado aparte con `IMPORTRANGE`.
- **Pegados de varias celdas.** Si alguien quita una protección y pega varias celdas, el script no puede restaurar los valores anteriores: lo registra, avisa y restaura las protecciones. La recuperación se hace con el historial de versiones.
- **Margen de segundos.** Entre que un profesor escribe en una fila libre y el script la protege pasan unos segundos.
- **Un solo instalador.** Los disparadores corren con la cuenta que ejecutó *Configurar libro*. Si esa cuenta se desactiva, otro ADMIN debe volver a ejecutarlo; primero hay que borrar los disparadores anteriores en Apps Script → *Activadores*.
