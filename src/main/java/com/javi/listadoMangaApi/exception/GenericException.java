package com.javi.listadoMangaApi.exception;

public class GenericException extends Exception {

    private static final long serialVersionUID = 7820084248357272099L;

    public GenericException(String mensaje) {
	super(mensaje);
    }
}
