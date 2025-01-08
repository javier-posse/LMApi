package com.javi.listadoMangaApi.config;

import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.env.Environment;

@Configuration
@PropertySource("classpath:errors.properties")
public class ErrorsConfig {

    private Environment env;

    public ErrorsConfig(Environment env) {
	this.env = env;
    }

    public String getProperty(String propertyKey) {
	return env.getProperty(propertyKey);
    }
}
