package com.javi.listadoMangaApi.dto;

import java.util.List;

import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class AuthorDto {
    @NotNull(message = "Id cannot be null")
    private int id;
    @NotNull(message = "Name cannot be null")
    private String name;
    private String biography;
    private List<SeriesSimplifiedDto> series;
}
