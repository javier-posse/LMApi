package com.javi.listadoMangaApi.services;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.javi.listadoMangaApi.commons.ExcelGenerator;
import com.javi.listadoMangaApi.dto.AuthorDto;
import com.javi.listadoMangaApi.dto.CollectionDto;
import com.javi.listadoMangaApi.dto.JpPublisherDto;
import com.javi.listadoMangaApi.dto.MonthReleasesDto;
import com.javi.listadoMangaApi.dto.SeriesDto;
import com.javi.listadoMangaApi.dto.SeriesReleaseDto;
import com.javi.listadoMangaApi.dto.SpPublisherDto;
import com.javi.listadoMangaApi.exception.GenericException;
import com.javi.listadoMangaApi.scrapers.AuthorScraper;
import com.javi.listadoMangaApi.scrapers.CollectionScraper;
import com.javi.listadoMangaApi.scrapers.JpPublisherScraper;
import com.javi.listadoMangaApi.scrapers.MonthReleasesScraper;
import com.javi.listadoMangaApi.scrapers.SeriesScraper;
import com.javi.listadoMangaApi.scrapers.SpPublisherScrapper;

@Service
public class ListadoMangaApiService {
    @Autowired
    AuthorScraper authorScraper;
    @Autowired
    CollectionScraper collectionScraper;
    @Autowired
    SpPublisherScrapper publisherScrapper;
    @Autowired
    JpPublisherScraper jpPublisherScraper;
    @Autowired
    SeriesScraper seriesScraper;
    @Autowired
    MonthReleasesScraper monthReleasesScraper;

    public AuthorDto searchAuthorById(int id) throws GenericException {
	return authorScraper.scrapAuthorPage(id);
    }

    public CollectionDto searchCollectionById(int id) throws GenericException {
	return collectionScraper.scrapCollectionPage(id);
    }

    public SpPublisherDto searchSpPublisherById(int id) throws GenericException {
	return publisherScrapper.scrapSpPublisherPage(id);
    }

    public JpPublisherDto searchJpPublisherById(int id) throws GenericException {
	return jpPublisherScraper.scrapJpPublisherPage(id);
    }

    public SeriesDto searchSeriesById(int id) throws GenericException {
	return seriesScraper.scrapSeriesPage(id);
    }

    public MonthReleasesDto searchMonthReleases(int month, int year) throws GenericException {
	return monthReleasesScraper.scrapMonthReleasesPage(month, year);
    }

    public MonthReleasesDto searchYearReleases(int year, boolean generateExcel) throws GenericException {
	MonthReleasesDto yearReleases = new MonthReleasesDto("Novedades " + year, new ArrayList<>());
	int lastVolumeCount = 0;
	for (int i = 1; i < 13; i++) {
	    // Creo un list para contar la cantidad de últimos tomos y luego lo piso
	    List<SeriesReleaseDto> seriesReleases = monthReleasesScraper.scrapMonthReleasesPage(i, year)
		    .getSeriesReleases();

	    // Uso el list anterior para filtrar en que importa de verdad
	    yearReleases.getSeriesReleases()
		    .addAll(seriesReleases.stream()
			    .filter(seriesRelease -> seriesRelease.isFirstRelease() || seriesRelease.isOnlyVolume())
			    .collect(Collectors.toList()));

	    lastVolumeCount += seriesReleases.stream().filter(seriesRelease -> seriesRelease.isLastVolume()).count();
	}
	if (generateExcel) {
	    ExcelGenerator.generateExcelReleases(year, yearReleases.getSeriesReleases(), lastVolumeCount);
	}

	return yearReleases;

    }
}
