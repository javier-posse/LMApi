package com.javi.listadoMangaApi.controllers;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.javi.listadoMangaApi.dto.AuthorDto;
import com.javi.listadoMangaApi.dto.CollectionDto;
import com.javi.listadoMangaApi.dto.JpPublisherDto;
import com.javi.listadoMangaApi.dto.MonthReleasesDto;
import com.javi.listadoMangaApi.dto.SeriesDto;
import com.javi.listadoMangaApi.dto.SpPublisherDto;
import com.javi.listadoMangaApi.exception.GenericException;
import com.javi.listadoMangaApi.services.ListadoMangaApiService;

@RestController
public class ListadoMangaApiControllerV1 {
    @Autowired
    ListadoMangaApiService apiService;

    @GetMapping("/author/{id}")
    public AuthorDto getAuthorInfo(@PathVariable int id) throws GenericException {
	return apiService.searchAuthorById(id);
    }

    @GetMapping("/collection/{id}")
    public CollectionDto getCollectionInfo(@PathVariable int id) throws GenericException {
	return apiService.searchCollectionById(id);
    }

    @GetMapping("/spPublisher/{id}")
    public SpPublisherDto getSpPublisherInfo(@PathVariable int id) throws GenericException {
	return apiService.searchSpPublisherById(id);
    }

    @GetMapping("/jpPublisher/{id}")
    public JpPublisherDto getJpPublisherInfo(@PathVariable int id) throws GenericException {
	return apiService.searchJpPublisherById(id);
    }

    @GetMapping("/series/{id}")
    public SeriesDto getSeriesInfo(@PathVariable int id) throws Exception {
	return apiService.searchSeriesById(id);
    }

    @GetMapping("/monthReleases/{month}/{year}")
    public MonthReleasesDto getMonthReleases(@PathVariable int month, @PathVariable int year) throws GenericException {
	return apiService.searchMonthReleases(month, year);
    }

    @GetMapping("/yearReleases/{year}")
    public MonthReleasesDto getMonthReleases(@PathVariable int year, @RequestParam boolean generateExcel)
	    throws GenericException {
	return apiService.searchYearReleases(year, generateExcel);
    }

}
