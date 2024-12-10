package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.net.MalformedURLException;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.CollectionDto;
import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class CollectionScraper {

    public CollectionDto scrapCollectionPage(int id) throws GenericException {
	String link = UrlConstants.BASE_URL + UrlConstants.PUBLISHER_COLLECTION + "?" + UrlConstants.ID_PARAM + "="
		+ id;
	// logger.info(link);
	CollectionDto collectionDto = null;

	try {
	    // get collection name
	    Document doc = Jsoup.connect(link).get();
	    String collectionName = doc.getElementsByClass("cen").select("h2").first().text();

	    // get series
	    List<SeriesSimplifiedDto> series = CommonUtils.getSeriesList(doc);

	    collectionDto = new CollectionDto(id, collectionName, series);

	} catch (MalformedURLException exception) {
	    throw new GenericException("Ha habido un error no controlado");
	} catch (IOException exception) {
	    throw new GenericException("Ha habido un error no controlado");
	}

	return collectionDto;
    }

}
