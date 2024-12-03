package com.javi.listadoMangaApi.commons;

import java.util.ArrayList;
import java.util.List;

import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;

public class CommonScrapers {

    public static List<SeriesSimplifiedDto> getSeriesList(Document doc) {
	List<SeriesSimplifiedDto> series = new ArrayList<>();
	Elements publisherSeries = doc.select("[class^=\"ventana_id\"]");

	publisherSeries.forEach(element -> {
	    Elements seriesNameElement = element.getElementsByTag("a");
	    seriesNameElement.forEach(seriesName -> {
		int seriesId = getLinkId(seriesName);
		String name = seriesName.text();
		SeriesSimplifiedDto serie = new SeriesSimplifiedDto(seriesId, name);
		series.add(serie);
	    });
	});

	return series;

    }

    public static int getLinkId(Element link) {
	return Integer.parseInt(link.attr("href").split("=")[1]);
    }

}
