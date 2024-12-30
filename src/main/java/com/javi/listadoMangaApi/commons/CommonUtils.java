package com.javi.listadoMangaApi.commons;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CommonUtils {

    private CommonUtils() {

    }

    private static final DateTimeFormatter FULL_DATE_FORMATTER = new DateTimeFormatterBuilder().parseCaseInsensitive()
	    .appendPattern("EEEE, d MMMM yyyy").toFormatter(Locale.forLanguageTag("es-ES"));

    private static final DateTimeFormatter YEAR_MONTH_FORMATTER = new DateTimeFormatterBuilder().parseCaseInsensitive()
	    .appendPattern("MMMM yyyy").toFormatter(Locale.forLanguageTag("es-ES"));

    public static List<SeriesSimplifiedDto> getSeriesList(Document doc) {
	List<SeriesSimplifiedDto> series = new ArrayList<>();
	Elements publisherSeries = doc.select("[class^=\"ventana_id\"]");

	publisherSeries.forEach(element -> {
	    Elements seriesNameElement = element.getElementsByTag("a");
	    seriesNameElement.forEach(seriesName -> {
		int seriesId = getIdFromLink(seriesName);
		String name = seriesName.text();
		SeriesSimplifiedDto serie = new SeriesSimplifiedDto(seriesId, name);
		series.add(serie);
	    });
	});

	return series;

    }

    public static int getIdFromLink(Element link) {
	return Integer.parseInt(link.attr("href").split("=")[1]);
    }

    public static String getLinkFromId(int id, String type) {
	StringBuilder linkSb = new StringBuilder(UrlConstants.BASE_URL);
	linkSb.append(type).append("?").append(UrlConstants.ID_PARAM).append("=").append(id);
	return linkSb.toString();
    }

    public static LocalDate fullDateConverter(String date) {
	if (date.contains(",")) {
	    return LocalDate.parse(date, FULL_DATE_FORMATTER);
	} else {
	    return YearMonth.parse(date, YEAR_MONTH_FORMATTER).atEndOfMonth();
	}
    }

}
