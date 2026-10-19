# Versión Excel (OneDrive / SharePoint)

Pestaña **Excel para OneDrive** de la app. Genera un solo .xlsx protegido con contraseñas que se sube a OneDrive o SharePoint.

**Qué se carga y qué se guarda** (siempre donde usted elija):

| | Botón | Archivo |
|---|---|---|
| Entrada | *Elegir archivos… / Elegir carpeta…* (arriba) | Plantillas .xlsx de los programas (o las incluidas en la app) |
| Entrada / salida | *Importar CSV… / Exportar CSV…* | Profesores con filas reservadas (`programa,profesor,filas`) |
| Entrada / salida | *Cargar archivo… / Guardar archivo…* | Contraseñas (`rol,ambito,contrasena`). Guárdelo en un lugar privado |
| Salida | *Generar Excel para OneDrive…* | El libro que se sube a OneDrive |

## Qué contiene el Excel

| Hoja | Contenido |
|---|---|
| **Instrucciones** | Explicación para usuarios (solo lectura) |
| **Una hoja por programa** (AFIN, EMBA, …) | Las columnas de las plantillas y las filas de registro elegidas (150 por defecto) |
| **Consolidado** | Une automáticamente todo lo registrado en las hojas de programa (solo lectura, requiere Excel 365 / Excel web) |

Se respeta el código de colores de las plantillas: **azul = Programa**, **verde = CASA**. Si a alguna plantilla le falta una columna que otras sí tienen (por ejemplo, EMBA no tenía "Nombre actividad"), se agrega para que todas las hojas sean iguales.

## Modelo de permisos

| Zona | Quién puede editar | Cómo |
|---|---|---|
| Títulos, leyenda, encabezados, columna "Programa", hoja Consolidado, Instrucciones | Nadie (solo administración) | Hoja protegida con contraseña de administrador |
| Agregar/borrar/renombrar hojas | Nadie (solo administración) | Estructura del libro protegida con la contraseña de administrador |
| Columnas **azules** de la hoja de un programa | Solo ese programa | Rango con la contraseña de ese programa |
| Filas asignadas a un profesor (azules) | Solo ese profesor | Rango con contraseña personal; el programa ya no puede editarlas y el nombre del profesor queda fijo |
| Columnas **verdes** de todas las hojas | Solo CASA | Rango con la contraseña del equipo CASA |

Cuando alguien intenta escribir en una celda de color, Excel le pide la contraseña del rango. Con la de AFIN no se puede editar MF, y con la de un profesor no se pueden editar las filas de otro.

### Asignar filas a profesores (requisito: lo que crea un profesor no lo modifica otro)

Excel no sabe qué persona está editando: solo controla con contraseñas. Por eso la propiedad de las filas se logra **reservando filas para cada profesor**, cada una con su propia contraseña:

1. En *1. Configuración*, agregue cada profesor (programa, nombre y cantidad de filas), o use *Importar CSV…* con un archivo `programa,profesor,filas`.
   Las filas se asignan en orden desde la fila 4 de la hoja del programa. Las filas sobrantes quedan para el programa (gestores).
2. Cargue el archivo de contraseñas que ya tenía (*Cargar archivo…*) y genere el Excel de nuevo. Las contraseñas existentes se conservan y solo se crean las de los profesores nuevos.

Si un programa no tiene profesores, toda su zona azul se edita con la contraseña del programa.

## Generar el archivo

1. Arriba, confirme las plantillas (por defecto, las incluidas en la app).
2. Si ya entregó contraseñas antes, use **Cargar archivo…** para reutilizarlas. Si no, se crean nuevas.
3. Clic en **Generar Excel para OneDrive…** y elija dónde guardar el libro.
4. Si se crearon o cambiaron contraseñas, la app le pide guardarlas. **Hágalo**: sin ese archivo no podrá volver a generar el Excel con las mismas claves. Si intenta cerrar la app con contraseñas sin guardar, también se lo recuerda.

⚠️ El Excel se genera **vacío**. Hágalo antes de empezar a registrar datos. Una vez el archivo esté en uso en OneDrive, los cambios de permisos se hacen a mano en Excel (*Revisar → Permitir modificar rangos* / *Administrar protección*) con la contraseña de administrador.

## Despliegue en OneDrive / SharePoint

1. Subir el Excel generado a una carpeta de OneDrive o SharePoint de CASA.
2. Compartir con **permiso de edición** a gestores, profesores y equipo CASA. (Si alguien solo debe consultar, se comparte como "Puede ver".)
3. Entregar a cada persona **solo su contraseña**, por un canal privado (correo o Teams directo), tomándola de la tabla *2. Contraseñas* (botón *Copiar seleccionada*) o del archivo de contraseñas:
   - cada programa recibe la suya, cada profesor la suya,
   - el equipo CASA recibe la de *Equipo CASA*,
   - la de *Administrador* solo la guarda la persona administradora.
4. Activar el control de cambios: en Excel web, **Revisar → Mostrar cambios** muestra quién modificó cada celda y cuándo. Junto con el **Historial de versiones** de OneDrive permite ver y revertir cualquier edición indebida.

## Limitaciones (importante)

- **La protección de Excel es un control de uso, no seguridad fuerte.** Evita errores y ediciones en la zona equivocada, pero una persona con conocimientos técnicos y permiso de edición podría saltársela, o compartir su contraseña. El respaldo real es *Mostrar cambios* + *Historial de versiones* (que registran el usuario real de Microsoft 365).
- La propiedad de filas por profesor exige saber de antemano cuántas filas necesita cada uno. Si un profesor necesita más, el administrador amplía su rango en *Administrar protección*.
- La hoja Consolidado usa `VSTACK`/`FILTER` (Excel 365 y Excel para la web). En Excel 2019 o anterior aparece como error.
- Probado en Excel para Mac: la protección de hojas y del libro queda activa y la hoja Consolidado reúne bien los datos de varios programas. **No se probó** en la interfaz que Excel pida la contraseña de cada rango, ni en Excel para la web; conviene hacer una prueba rápida con dos o tres usuarios antes de abrirlo a todos.

## Si se necesita control por usuario real

Si hace falta que el sistema reconozca a cada persona sin usar contraseñas:

- **Microsoft 365 (sin salir de OneDrive):** un archivo por programa, cada uno en una carpeta de SharePoint con permisos solo para ese programa y CASA, y un libro maestro de solo lectura que los une con Power Query. Así cada programa queda aislado por su cuenta institucional. Para filas por profesor, la herramienta indicada es **Microsoft Lists** con el permiso de lista "Crear elementos y editar solo los elementos creados por el usuario".
- **Google Sheets:** es la otra pestaña de esta app. Los permisos van por correo y un Apps Script bloquea cada fila para quien la creó.
