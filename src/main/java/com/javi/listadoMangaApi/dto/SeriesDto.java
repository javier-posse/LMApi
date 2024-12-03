package com.javi.listadoMangaApi.dto;

import java.util.List;

import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.ToString;

@AllArgsConstructor
@ToString
@Getter
public class SeriesDto {
    @NotNull(message = "Id cannot be null")
    private int id;
    @NotNull(message = "Name cannot be null")
    private String name;
    private String originalName;
    private String originalEdition;
    private int scripterId;
    private String scripterName;
    private int artistId;
    private String artistName;
    private int originalStoryId;
    private String originalStoryName;
    private int jpPublisherId;
    private String jpPublisherName;
    private int spPublisherId;
    private String spPublisherName;
    private int collectionId;
    private String collectionName;
    private String formFactor;
    private String reedOrder;
    private String jpPublishedVolumes;
    private String spPublishedVolumes;
    private List<VolumeDto> volumes;
}
