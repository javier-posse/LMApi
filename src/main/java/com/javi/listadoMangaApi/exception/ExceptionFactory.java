package com.javi.listadoMangaApi.exception;

import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.config.AppConfig;

@Component
public class ExceptionFactory {

    private final AppConfig configUtil;

    public ExceptionFactory(AppConfig configUtil) {
	this.configUtil = configUtil;
    }

    public GenericException createGenericException() {
	String message = configUtil.getProperty("error.generic.message");
	String code = configUtil.getProperty("error.generic.code");
	return new GenericException(message, code);
    }

    public ExcelException createExcelException() {
	String message = configUtil.getProperty("error.excel.message");
	String code = configUtil.getProperty("error.excel.code");
	return new ExcelException(message, code);
    }
}