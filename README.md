# coordsFromGeoportaltoExcel
VBA script that allows to obtain and upload to Excel the coordinates of the client of the CNEL EP geoportal (Ecuador) from the unique national code of the client through web requests and json responses.


Execution Test:

![Imgur](https://i.imgur.com/QwJ3mu7.gif)

To execute:
* Visit the geoportal and perform any search by unique code:

https://geoportal.cnelep.gob.ec/cnel/


![Imgur1](https://i.imgur.com/MI9od5K.png)


![Imgur2](https://i.imgur.com/F7tfvMA.png)


* Access the development tools of our browser before executing the search (I usually use the f12 key in chrome) and press apply, we locate the first requested Query request in the network list of the development tools and display the details.

![Imgur4](https://i.imgur.com/UjTQend.png)


![Imgur3](https://i.imgur.com/H8Xr3QO.png)


* Create an excel file with a specific sheet from which the code will obtain the references of which columns contain coordinates (except for the first row).
* Copy the selection (url before the word query) and transfer it to our Excel VAR sheet:

VAR Sheet:


![Imgur4](https://i.imgur.com/BQ1qaDC.png)


![Imgur5](https://i.imgur.com/vpPBbRI.png)


* Import the .bas .cls modules from the excel VBA editor.
Special thanks to the post https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel


![Imgur6](https://i.imgur.com/aSbpjgJ.png)


* In Excel, build the following table on an empty sheet, paying special attention to the columns specified in the VAR sheet in the previous step. The columns must match the headers, not textually, but they must be the data specified in the VAR sheet.


![Imgur7](https://i.imgur.com/xQoRmda.png)


* Execute the macro according to the need and requirement.

* Once the table has data, it can be executed by selecting one or more elements from the UNIC CODE column (column A), as long as there is reference data to perform the search.


![Imgur8](https://i.imgur.com/QwJ3mu7.gif)


Bibliography:

https://www.codeproject.com/Articles/828911/Recursive-VBA-JSON-Parser-for-Excel