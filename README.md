# API Listado Manga
Este repositorio contiene un proyecto desarrollado en Java Spring que actúa como API para la web **Listado Manga** mediante técnicas de web scraping. 
## Descripción
El objetivo de este proyecto es proporcionar una API que permita acceder y consultar la información disponible en la web Listado Manga. 
Esta API es fruto del cariño y admiración que tengo por la web, y reconocimiento a su duro trabajo. 
La subo con el beneplácito de Listado Manga.

## Funcionalidades
- Obtener información de autor.
- Obtener información de colección editorial.
- Obtener información de editorial española.
- Obtener información de editorial japonesa.
- Obtener información de una serie.
- Obtener lanzamientos de un mes.
- Obtener información de las novedades de un año.
- Generar un excel con la información de novedades del año y algunas estadísticas.

## Tecnologías Utilizadas 
- **Java Spring**: Framework principal del proyecto.
- **Jsoup**: Librería para el web scraping.
- **Lombok**: Librería de limpieza de código.
- **Apache POI**: Librería para generación de ficheros office. Se usa para generar un excel.
- **Maven**: Gestión de librerías.

## Licencia
Para más detalles, consulta el archivo LICENSE.

## Uso
Una vez ejecutado como se prefiera, se pueden encontrar las llamadas de postman en /src/main/resources/Postman_Collections.

El excel generado se puede encontrar en /Excel.
Por favor, tenga en cuenta que el código no tendrá ningún excel ya generado de ningún año. Si quiere uno, debe lanzar el endpoint de novedades de un año con el parámetro correspondiente.

Hay un dockerfile. Si se va a usar docker para desplegarlo, hay que montar el path de /Excel para poder acceder a él. Se crea en el root del docker.

## Errores conocidos
La funcionalidad de la serie debe ser reescrita casi por completo.
