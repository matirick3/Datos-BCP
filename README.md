Subí acá el boletín mensual del BCP tal cual lo bajás de su web (archivo `.xlsm`, el que
trae la hoja "5. Cred. por sector").

En cuanto lo subís (Add file → Upload files, directo a esta carpeta), una GitHub Action se
dispara sola, extrae los datos y actualiza `data/creditos-sector.json` — la web los toma de
ahí automáticamente. El archivo `.xlsm` que subiste se borra solo una vez procesado (ya
quedó extraído en el JSON, no hace falta guardarlo acá).

El sitio del BCP bloquea las descargas automatizadas, así que este paso de subir el archivo
tiene que hacerlo una persona — no hay forma de automatizarlo del todo.
