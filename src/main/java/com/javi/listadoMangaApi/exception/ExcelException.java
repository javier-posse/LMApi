package com.javi.listadoMangaApi.exception;

public class ExcelException extends CustomException {

    private static final long serialVersionUID = 7820084248357272099L;

    public ExcelException(String mensaje, String code) {
	super(mensaje, code);
    }
}
