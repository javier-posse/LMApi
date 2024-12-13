package com.javi.listadoMangaApi.exception;

public class GenericException extends CustomException {

    private static final long serialVersionUID = 7820084248357272099L;

    public GenericException(String mensaje, String code) {
	super(mensaje, code);
    }
}
