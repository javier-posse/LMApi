FROM openjdk:21
ADD target/listadoMangaApi-1.0.0-SNAPSHOT.jar app1.jar
ENTRYPOINT [ "java", "-jar","app1.jar" ] 