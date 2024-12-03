package com.javi.listadoMangaApi.dto;

import java.util.List;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class MonthReleasesDto {
    private String monthName;
    private List<SeriesReleaseDto> seriesReleases;
}
