package com.javi.listadoMangaApi.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import lombok.ToString;

@AllArgsConstructor
@NoArgsConstructor
@ToString
@Getter
@Setter
public class SeriesInfoDto {
    private String originalName;
    private String originalEdition;
    private AuthorDto[] authors;
    private int scripterId;
    private String scripterName;
    // También se usa para "ilustraciones" en novelas.
    private int artistId;
    private String artistName;
    private int originalStoryId;
    private String originalStoryName;
    private int characterDesignId;
    private String characterDesignName;
    private int colorId;
    private String colorName;
    private int translatorId;
    private String translatorName;
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
    private String notes;
}
