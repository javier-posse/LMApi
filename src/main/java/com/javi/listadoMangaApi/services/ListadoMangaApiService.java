package com.javi.listadoMangaApi.services;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.springframework.stereotype.Service;

import com.javi.listadoMangaApi.commons.ExcelGenerator;
import com.javi.listadoMangaApi.dto.AuthorDto;
import com.javi.listadoMangaApi.dto.CollectionDto;
import com.javi.listadoMangaApi.dto.JpPublisherDto;
import com.javi.listadoMangaApi.dto.MonthReleasesDto;
import com.javi.listadoMangaApi.dto.SeriesDto;
import com.javi.listadoMangaApi.dto.SeriesReleaseDto;
import com.javi.listadoMangaApi.dto.SpPublisherDto;
import com.javi.listadoMangaApi.exception.ExceptionFactory;
import com.javi.listadoMangaApi.exception.GenericException;
import com.javi.listadoMangaApi.scrapers.AuthorScraper;
import com.javi.listadoMangaApi.scrapers.CollectionScraper;
import com.javi.listadoMangaApi.scrapers.JpPublisherScraper;
import com.javi.listadoMangaApi.scrapers.MonthReleasesScraper;
import com.javi.listadoMangaApi.scrapers.SeriesScraper;
import com.javi.listadoMangaApi.scrapers.SpPublisherScrapper;

@Service
public class ListadoMangaApiService {

    AuthorScraper authorScraper;
    CollectionScraper collectionScraper;
    SpPublisherScrapper publisherScrapper;
    JpPublisherScraper jpPublisherScraper;
    SeriesScraper seriesScraper;
    MonthReleasesScraper monthReleasesScraper;
    ExceptionFactory exceptionFactory;

    public ListadoMangaApiService(AuthorScraper authorScraper, CollectionScraper collectionScraper,
	    SpPublisherScrapper publisherScrapper, JpPublisherScraper jpPublisherScraper, SeriesScraper seriesScraper,
	    MonthReleasesScraper monthReleasesScraper, ExceptionFactory exceptionFactory) {
	this.authorScraper = authorScraper;
	this.collectionScraper = collectionScraper;
	this.publisherScrapper = publisherScrapper;
	this.jpPublisherScraper = jpPublisherScraper;
	this.seriesScraper = seriesScraper;
	this.monthReleasesScraper = monthReleasesScraper;
	this.exceptionFactory = exceptionFactory;
    }

    public AuthorDto searchAuthorById(int id) throws GenericException {
	try {
	    return authorScraper.scrapAuthorPage(id);
	} catch (IOException e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public CollectionDto searchCollectionById(int id) throws GenericException {
	try {
	    return collectionScraper.scrapCollectionPage(id);
	} catch (IOException e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public SpPublisherDto searchSpPublisherById(int id) throws GenericException {
	try {
	    return publisherScrapper.scrapSpPublisherPage(id);
	} catch (IOException e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public JpPublisherDto searchJpPublisherById(int id) throws GenericException {
	try {
	    return jpPublisherScraper.scrapJpPublisherPage(id);
	} catch (IOException e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public SeriesDto searchSeriesById(int id) throws GenericException {
	try {
	    return seriesScraper.scrapSeriesPage(id);
	} catch (Exception e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public MonthReleasesDto searchMonthReleases(int month, int year) throws GenericException {
	try {
	    return monthReleasesScraper.scrapMonthReleasesPage(month, year);
	} catch (IOException e) {
	    throw exceptionFactory.createGenericException();
	}
    }

    public MonthReleasesDto searchYearReleases(int year, boolean generateExcel) throws GenericException {
	MonthReleasesDto yearReleases = new MonthReleasesDto("Novedades " + year, new ArrayList<>());
	int lastVolumeCount = 0;
	for (int i = 1; i < 13; i++) {
	    // Creo un list para contar la cantidad de últimos tomos y luego lo piso
	    List<SeriesReleaseDto> seriesReleases;
	    try {
		seriesReleases = monthReleasesScraper.scrapMonthReleasesPage(i, year).getSeriesReleases();
	    } catch (IOException e) {
		throw exceptionFactory.createGenericException();
	    }

	    // Uso el list anterior para filtrar en que importa de verdad
	    yearReleases.getSeriesReleases().addAll(seriesReleases.stream()
		    .filter(seriesRelease -> seriesRelease.isFirstRelease() || seriesRelease.isOnlyVolume()).toList());

	    lastVolumeCount += seriesReleases.stream().filter(SeriesReleaseDto::isLastVolume).count();
	}
	if (generateExcel) {
	    try {
		ExcelGenerator.generateExcelReleases(year, yearReleases.getSeriesReleases(), lastVolumeCount);
	    } catch (IOException e) {
		throw exceptionFactory.createGenericException();
	    }
	}

	return yearReleases;

    }
}
