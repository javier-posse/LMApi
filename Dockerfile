FROM openjdk:21
ADD target/listadoMangaApi-1.0.2-SNAPSHOT.jar lmapi.jar
RUN mkdir /Excel
ENTRYPOINT [ "java", "-jar","lmapi.jar" ] 