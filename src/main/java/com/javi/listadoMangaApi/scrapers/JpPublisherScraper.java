package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.JpPublisherDto;
import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;
import com.javi.listadoMangaApi.dto.SpPublisherDto;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class JpPublisherScraper {

    public JpPublisherDto scrapJpPublisherPage(int id) throws IOException {
	log.info("JpPublisherScraper inicializado");
	String link = UrlConstants.BASE_URL + UrlConstants.ORIGINAL_PUBLISHER_PATH + "?" + UrlConstants.ID_PARAM + "="
		+ id;
	JpPublisherDto jpPublisherDto = null;

	// get publisher name
	Document doc = Jsoup.connect(link).timeout(10000).get();
	Elements allPublishers = doc.getElementsByClass("cen").select("h2");
	String publisherName = allPublishers.first().text();
	int nameEnd = publisherName.indexOf(" editadas");
	publisherName = publisherName.substring(15, nameEnd);
	Elements allSquares = doc.select("[class^=\"ventana_id\"]");
	List<SpPublisherDto> spPublishers = new ArrayList<>();
	String spPublisherName = null;
	boolean isFirst = true;
	List<SeriesSimplifiedDto> series = new ArrayList<>();

	for (Element square : allSquares) {
	    Elements seriesElem = square.select("a");
	    if (seriesElem.isEmpty()) {
		int nameStart = square.text().indexOf("por ");
		spPublisherName = square.text().substring(nameStart + 4, square.text().length());

		if (isFirst) {
		    isFirst = false;
		} else {
		    SpPublisherDto spPublisher = new SpPublisherDto(0, spPublisherName, series);
		    spPublishers.add(spPublisher);
		    series = new ArrayList<>();
		}
	    } else {
		int seriesId = CommonUtils.getIdFromLink(seriesElem.first());
		String name = seriesElem.text();
		SeriesSimplifiedDto serie = new SeriesSimplifiedDto(seriesId, name);
		series.add(serie);
	    }

	}
	// get series
	jpPublisherDto = new JpPublisherDto(id, publisherName, spPublishers);

	return jpPublisherDto;
    }

}
