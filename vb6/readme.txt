Attribute VB_Name = "readme"
'-------------------------------------------------------------------
' Notas de la version 1.1
' Laurens Rodriguez
'-------------------------------------------------------------------

'   27/08/03:
'---------------
'   He agregado la fecha del archivo a la exploracion (era necesaria la funcion: FileTimeToLocalFileTime)
'   Ademas esta la conocida diferencia de poner los datos como #12/15/2003# para access que es diferente en SQL Server (se pone solo '12/15/2003')
'   Ah un dato mas... la tabla FILE tenia un campo DATE que no me hizo problemas hasta que quize usar INSERT ...
'   al parecer no acepta una tabla con este nombre por eso la cambie a FECHA.
'   Hasta el momento se ha culminado:
'   - Encriptacion de cadenas importantes (clase Crypto sin periodicidad)
'   - Inclusion y expansion de las peliculas flash dentro del ejecutable

'   28/08/03:
'---------------
'   Agregado prioridad por medio de graficos
'   El estilo XP tiene distintos margenes del estilo clasico windows he mejorado en algo el evento resize del frmDataControl
'   Agregue un cmboption a las opciones de busqueda adicionales, se agrego busqueda por prioridad

'   29/08/03:
'---------------
'   Creado el control ActiveX ATL para extraer informacion del header MPEG Audio (VC++)
'   Agregada la opcion de info MPEG y bitrate cuando son MP3
'   Agregada la opcion de prioridad predeterminada

'   30/08/03:
'---------------
'   Agregada nueva interfaz al control ATL

'   02/09/03:
'---------------
'   Agregada opciones para extraccion del header MPEG (VBR y buffer)
'   debe ser mas rapido (pero menos exacto) a menor buffer y sin la opcion VBR

'   03/09/03:
'---------------
'   Busqueda por calidad y prioridad, corregido problema con el menu [Ver detalle]

'   11/09/03:
'---------------
'   Extraccion de los DSN del sistema

'   12/09/03:
'---------------
'   AddIn para generacion de codigo en avance - analisis de la DB completado:

'   13/09/03:
'---------------
'   Carga para edicion mejorada (soporta registros dañados). Opcion de guardar lo editado terminada.
'   Cambiada estructura DB:
'      Parent
'      ------------------------------------------------
'        parent              adVarWChar         [50]
'   En todas las tablas relacionadas el indice cero se usa para indicar registro vacio o no asignado
'   y SIEMPRE debe encontrarse.

'   17/09/03:
'---------------
'   Cambiada estructura DB:
'      File
'      ------------------------------------------------
'        type_comment        adVarWChar         [20]
'   Corregida descarga de formularios de edicion con: frmEditRegistros.Show vbModeless, frmDataControl

'   18/09/03:
'---------------
'   Mejorada descarga de formularios (con el anterior los formularios ocultaban el frmDataControl)
'   ahora se hace en el evento unload del frmDataControl
'   Agregada opcion de eliminar registros.
'   Corregido algunos defectillos de los controles OCX (fallaba en los bordes superior y izq. asi que cambie un " <= 0" por un " < 0")
'   Pero al volver a compilar esto impidio que el proyecto se cargara: No se reconocia el OCX! (y nada de registrarla usando regsvr32)
'   Solucion: si no quieres volver a dibujar tus controles haz un proyecto dummy y copia la declaracion del vbp:
'   Object={DD74340F-8899-4D8E-ABDC-577BDB507BFD}#1.0#0; exlightbutton.ocx a tu proyecto (que seguro tiene otro) y ya esta.

'   19/09/03:
'---------------
'   Avances con el add-in: ya se inserta un formulario plantilla al proyecto

'   20/09/03:
'---------------
'   El MSHFlexdrid 6.0 tiene un defecto con la opcion de enfoque ligero de celda, me di cuenta cuando agregaba texto
'   con el ForeColor cambiado (dejaba de ser lineas punteadas y se convertian en lineas completas), lo que da la impresion
'   de que algo anda mal, para que volviera a estar normal tenia que agregar texto en otra celda con el color default.
'   Asi que despues de tratar por un rato me dio colera y le puse enfoque ninguno (MSHFlexdrid de frmGenre).
'   Agregada la opcion de Ordenamiento (.Sort). Corregido un problema con eliminar registros nuevos en frmDataControl

'   22/02/04:
'---------------
'   He retomado el proyecto. Algunos cambios a [frmOpcBusqueda] creando proyecto vss

'   11/03/04:
'---------------
'   Arregle error (estetico) de resultados vacio y agregue combo de seleccion de tipo de medio de almacenamiento a frmDB
'   Comenze con quitar todos los << Me. >> del codigo (optimizar) primero comenze con frmDB

'   12/03/04:
'---------------
'   Quitada referencia  << Me. >> de frmOpcBusqueda (lo hago con reemplazar, sensible a MAYUS asi que espero no introducir errores indeseados)
'   agregado variable [gs_DBRegParent] para manejar busqueda por Pertenencia a Pariente y no solo a Medio. Esta busqueda ya funciona.
'   He modificado codigo para actualizar combos de frmDataControl y para mostrar los titulos del flexgrid de resultados.

'   13/03/04:
'---------------
'   Probando filtrado por tamaño. tipo y fecha (esta falta afinar para el caso igual a)

'   14/03/04:
'---------------
'   Filtrado por Fecha terminado y probado. Ahora falta la opcion de ordenar. Tambien agregue la
'   opcion de abrir/jecutar archivo con ShellExecute()

'   17/03/04:
'---------------
'   Agregada la opcion de mostrar en los resultados el Tipo de archivo o el Genero.

'   20/03/04:
'---------------
'   Corregido pequeño defectillo en frmDataControl (cuando resultados esta vacio y se ejecuta archivo)
'   Probado opcion de Ordenar (se agrego nuevo formulario para esto y un boton en frmDataControl)
'   Agregada capacidad para borrar multiples archivos de la BD
'   Se agrego la opcion de eliminar resultados de la lista (multiple seleccion)
'   corregi algunos errores de exploracion frmExplorar (se estaban omitiendo archivos)

'   21/03/04:
'---------------
'   Intente instalar en w98 y... se murio mi sistema w98 (remplazo dlls del sistema por dlls unicode de XP
'   no se arreglo con nada el msvcrt.dll lo pude recuperar pero habia un atl.dll y otros que no => mala idea)
'   He cargado el proyecto en W98 y me he dado cuenta de algunos defectos que ahora estoy tratando de arreglar)
'   Corregido Reporte en Win98, problema con el frmDB cuando no hay conexion, nuevo metodo de tratar archivos (hiden, readonly o system)

'   24/03/04:
'---------------
'   Detecte un problema en frmDataControl cuando se edita no destruye el objeto formulario de edicion
'   y despues de algunas ediciones me quedo sin memoria...
'   Con respecto al instalador pienso hacerlo en win98 y espero que corra en winXp
'   Los controles los he puesto en [compatibilidad del proyecto] para no estar cambiando cada vez los ID externamente

'   26/03/04:
'---------------
'   Arregle el error de falta de memoria cuando se edita y se agrega nuevos registros...

'   02/04/04:
'---------------
'   Agregado codigo de seguridad si el usuario ha cerrado la conexion y quiere hacer una consulta...


'   15/04/04:
'---------------
'   Agregado formulario que muestra la estructura de la base de datos, agregando formulario para ejecutar comandos SQL...

'   18/04/04:
'---------------
'   Se agrego la posibilidad de ejecutar un comando SQL, se cambio la forma de conectarse a la BD
'   para el instalador es necesario el atl.dll para W98 y el msstdfmt.dll (el vb5db.dll no es necesario)

'   24/04/04:
'---------------
'   Se agrego una tabla mas CATEGORY la cual sirve para clasificar los medios, generos y grupos de pertenencia
'
'   La clase encriptadora tenia un error cuando el codigo era por ejemplo 2004 (algoritmo para evitar periodicidad)
'
'   26/04/04:
'---------------
'   Se termino de integrar la nueva tabla en el programa. Pienso agregar la opcion de [Copiar a..] para mover los archivos del CD
'   Se dejo el mantenimiento de las otras tablas para la version 1.2 jaja. Estoy pensando en el soporte de script...hmmm
'
'   02/05/04:
'---------------
'   El formulario de SQL ahora permite visualizar los resultados de consultas SELECT ademas de exportar a EXCEL y texto
'
'   08/05/04:
'---------------
'   Agregue el control para poner colores de sintaxis SQL, me simplifica muchas cosas (esta vacan)
'   Estoy agregando la opcion de anexar registros a medios ya existentes
'
'   30/05/04:
'---------------
'   Corregido un big error cuando se guardaba el nombre del archivo en la BD (no se estaba controlando el tamaño por un
'   error de copy&paste).. ademas agregada la posibilidad de omitir los mensajes de recorte de campo y que en el campo nombre
'   se guarde el nombre del archivo sin tipo por defecto.
'   Modificada la tabla
'
'    **************************************************
'      FILE
'      ------------------------------------------------
'        name                adVarWChar         [85]
'
'   04/06/04:
'---------------
'   Se añadio la posibilidad de agregar directorios a la base de datos (ya no solo archivos)
'   Se corrigio un error de Mostrar_Detalle de frmDataControl.
'   agregada la capidad de exportacion al formulario frmDataControl usando el modulo mdlFlexExport
'
'   13/06/04:
'---------------
'   Trabajando en la tabla dinamica.
'   Se corrigio un error cuando en el frmDataControl se buscaba una cadena con apostrofe

'   16/08/04:
'---------------
'   Centralice las opciones en el formulario frmOptions.

'   21/09/04:
'---------------
'   Corregio error cuando el nombre del medio tenia comillas. Agregada la opcion de Salvar y Guardar scripts.

'   27/11/04:
'---------------
'   Agregue la opcion de usar plugins desde scripts

'   24/03/05:
'---------------
'   Por si me olvido el pwd del VSS: msk

'   23/05/05:
'---------------
'   agregada la posibilidad de buscar por nombre de archivo
'
'   agregado el campo para archivos ocultos
'
'    **************************************************
'      FILE
'      ------------------------------------------------
'        hidden                adVarWChar         [85]
'
'   29/07/05:
'---------------
'   quiero terminar con este programa ya!
'
'    **************************************************
'      Author
'      ------------------------------------------------
'        id_author           adInteger           4
'        id_category         adInteger           4
'        author              adVarWChar         [50]
'        active              adUnsignedTinyInt   1
'    **************************************************
'      category
'      ------------------------------------------------
'        id_category         adInteger           4
'        category            adVarWChar         [50]
'    **************************************************
'      FILE
'      ------------------------------------------------
'        id_file             adInteger           4
'        id_storage          adInteger           4
'        sys_name            adVarWChar         [255]
'        id_file_type        adInteger           4
'        sys_length          adInteger           4
'        name                adVarWChar         [75]
'        hidden              adUnsignedTinyInt   1
'        id_parent           adInteger           4
'        id_sys_parent       adInteger           4
'        id_author           adInteger           4
'        id_genre            adInteger           4
'        id_sub_genre        adInteger           4
'        fecha               adDBTimeStamp       16
'        priority            adUnsignedTinyInt   1
'        quality             adDouble            8
'        type_quality        adVarWChar         [8]
'        comment             adVarWChar         [50]
'        type_comment        adVarWChar         [20]
'    **************************************************
'      file_type
'      ------------------------------------------------
'        id_file_type        adInteger           4
'        file_type           adVarWChar         [15]
'    **************************************************
'      genre
'      ------------------------------------------------
'        id_genre            adInteger           4
'        id_category         adInteger           4
'        genre               adVarWChar         [30]
'        active              adUnsignedTinyInt   1
'    **************************************************
'      Parent
'      ------------------------------------------------
'        id_parent           adInteger           4
'        id_category         adInteger           4
'        parent              adVarWChar         [50]
'    **************************************************
'      R_porter
'      ------------------------------------------------
'        version             adVarWChar         [5]
'        autor               adVarWChar         [30]
'    **************************************************
'      STORAGE
'      ------------------------------------------------
'        id_storage          adInteger           4
'        id_storage_type     adInteger           4
'        id_category         adInteger           4
'        name                adVarWChar         [25]
'        label               adVarWChar         [20]
'        serial              adVarWChar         [10]
'        fecha               adDBTimeStamp       16
'        comment             adVarWChar         [50]
'        active              adUnsignedTinyInt   1
'    **************************************************
'      storage_type
'      ------------------------------------------------
'        id_storage_type     adInteger           4
'        storage_type        adVarWChar         [12]
'    **************************************************
'      sub_genre
'      ------------------------------------------------
'        id_sub_genre        adInteger           4
'        sub_genre           adVarWChar         [20]
'        id_genre            adInteger           4
'
'   01/10/06:
'---------------
'   1.2: it's finished I think
'   - the blody dynamic edition of tables is complete (now the edition cycle is closed)
'   - the DB explorer is now part of the program not longer a plugin (hell useful)
'   - the jscript could be handy sometimes apart of that the plugin and vb script
'     support are just a nice *this can be done too* but now I would use a simple
'     Ruby or Perl script instead of traumatizing my brain writing vbScript code :)
'   TODO (if hundreds of fanatic users were requesting it)
'   - Localization (english, English, ENGLISH)
'   - Port the damn thing to C++ using (Qt + sqlite + lua) and make it multiplatform
'   - Modern options dialog
'   - Skinning of main

