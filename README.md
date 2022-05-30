# coordsFromGeoportaltoExcel
Script VBA que permite obtener y subir a Excel las coordenadas del cliente del geoportal CNEL EP (Ecuador) a partir del código nacional único del cliente mediante solicitudes web y respuestas json.


Prueba de Ejecucion:


![Imgur](https://i.imgur.com/QwJ3mu7.gif)


Para ejecutar:

*Visite el geoportal y realice una busqueda cualquiera por codigo unico:
https://geoportal.cnelep.gob.ec/cnel/


![Imgur1](https://i.imgur.com/MI9od5K.png)


![Imgur2](https://i.imgur.com/F7tfvMA.png)


* Acceda a las herramientas de desarrollo de nuestro navegador antes de ejecutar la busqueda (generalmente utilizo la tecla f12 en chrome) y presionamos aplicar, ubicamos en la lista de red de las herramientas de desarrollo la primera peticion Query solicitada y desplegamos los detalles.


![Imgur4](https://i.imgur.com/UjTQend.png)


![Imgur3](https://i.imgur.com/H8Xr3QO.png)


* Cree un archivo excel con una hoja especifica de la cual el codigo obtendra las referencias de cuales son las columnas que contienen coordenadas(a excepcion de la primera fila).
Copiamos lo seleccionado (url antes de la palabra query) y lo trasladamos a nuestra hoja VAR de Excel:

Hoja VAR:


![Imgur4](https://i.imgur.com/BQ1qaDC.png)


![Imgur5](https://i.imgur.com/vpPBbRI.png)


* Importe los modulos .bas .cls desde el editor de VBA de excel.
Agadecimiento especial al post https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel


![Imgur6](https://i.imgur.com/aSbpjgJ.png)


* En Excel construya en una hoja vacia la siguiente tabla poniendo especial atencion a las columnas especificadas en la hoja VAR en el paso anterior las columnas deben concordar con los encabezados, no textualmente pero si deben ser los datos que se especificaron el la hoja VAR.


![Imgur7](https://i.imgur.com/xQoRmda.png)


* Ejecutar la macro segun la necesidad y requerimiento.

Una vez la tabla tenga datos se puede ejecutar seleccionando uno a varios elementos de la columna UNIC CODE (columna A), esto siempre que haya datos de referencia para realizar la busqueda.


![Imgur8](https://i.imgur.com/QwJ3mu7.gif)


Bibliografia:

https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel