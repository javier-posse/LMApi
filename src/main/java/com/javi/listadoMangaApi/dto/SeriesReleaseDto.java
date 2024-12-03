package com.javi.listadoMangaApi.dto;

import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class SeriesReleaseDto {
    @NotNull(message = "Id cannot be null")
    private int id;
    @NotNull(message = "Name cannot be null")
    private String name;
    private String releaseDate;
    private int authorId;
    private String authorName;
    private int artistId;
    private String artistName;
    private boolean lastVolume;
    private boolean firstRelease;
    private boolean onlyVolume;
}
