package com.javi.listadoMangaApi.exception;

public abstract class CustomException extends Exception {

    private static final long serialVersionUID = -3738401760065525929L;
    private final String errorCode;

    protected CustomException(String message, String errorCode) {
	super(message);
	this.errorCode = errorCode;
    }

    public String getErrorCode() {
	return errorCode;
    }
}