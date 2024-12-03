package com.javi.listadoMangaApi.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class VolumeDto {
    private String name;
    private String pages;
    private String price;
    private String releaseDate;
}
