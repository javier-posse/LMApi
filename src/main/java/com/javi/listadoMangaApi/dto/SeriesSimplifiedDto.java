package com.javi.listadoMangaApi.dto;

import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class SeriesSimplifiedDto {
    @NotNull(message = "Id cannot be null")
    private int id;
    @NotNull(message = "Name cannot be null")
    private String name;
}
