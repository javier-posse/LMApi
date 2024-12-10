package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.net.MalformedURLException;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;
import com.javi.listadoMangaApi.dto.SpPublisherDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class SpPublisherScrapper {

    public SpPublisherDto scrapSpPublisherPage(int id) throws GenericException {
	String link = UrlConstants.BASE_URL + UrlConstants.PUBLISHER_PATH + "?" + UrlConstants.ID_PARAM + "=" + id;
	SpPublisherDto spPublisherDto = null;

	try {
	    // get publisher name
	    Document doc = Jsoup.connect(link).get();
	    String publisherName = doc.getElementsByClass("cen").select("h2").first().text();
	    log.info(publisherName);

	    // get series
	    List<SeriesSimplifiedDto> series = CommonUtils.getSeriesList(doc);

	    spPublisherDto = new SpPublisherDto(id, publisherName, series);

	} catch (MalformedURLException exception) {
	    throw new GenericException("Ha habido un error no controlado");
	} catch (IOException exception) {
	    throw new GenericException("Ha habido un error no controlado");
	}

	return spPublisherDto;
    }

}
