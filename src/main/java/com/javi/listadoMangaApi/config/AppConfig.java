package com.javi.listadoMangaApi.config;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.env.Environment;

@Configuration
@PropertySource("classpath:errors.properties")
public class AppConfig {

    @Autowired
    private Environment env;

    public String getProperty(String pPropertyKey) {
	return env.getProperty(pPropertyKey);
    }
}
