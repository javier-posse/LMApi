package com.javi.listadoMangaApi.config;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelConfig {
    private Properties properties;

    public ExcelConfig() {
	properties = new Properties();
	try (InputStream input = getClass().getClassLoader().getResourceAsStream("excel.properties")) {
	    if (input == null) {
		log.error("No se ha encontrado fichero de prop");
		return;
	    }
	    properties.load(input);
	} catch (IOException ex) {
	    ex.printStackTrace();
	}
    }

    public String getProperty(String key) {
	return properties.getProperty(key);
    }
}