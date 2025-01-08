package com.javi.listadoMangaApi.exception;

import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.config.ErrorsConfig;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ExceptionFactory {

    private final ErrorsConfig configUtil;

    public ExceptionFactory(ErrorsConfig configUtil) {
	this.configUtil = configUtil;
    }

    public GenericException createGenericException(Exception e) {
	String message = configUtil.getProperty("error.generic.message");
	String code = configUtil.getProperty("error.generic.code");
	log.error(e.getMessage());
	return new GenericException(message, code);
    }

    public ExcelException createExcelException(Exception e) {
	String message = configUtil.getProperty("error.excel.message");
	String code = configUtil.getProperty("error.excel.code");
	log.error(e.getMessage());
	return new ExcelException(message, code);
    }
}