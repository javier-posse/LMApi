package com.javi.listadoMangaApi.commons;

import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeFormatterBuilder;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class CommonUtils {

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

    public static LocalDate fullDateConverter(String date) throws GenericException {
	try {
	    if (date.contains(",")) {
		return LocalDate.parse(date, FULL_DATE_FORMATTER);
	    } else {
		return YearMonth.parse(date, YEAR_MONTH_FORMATTER).atDay(1);
	    }
	} catch (DateTimeParseException e) {
	    throw new GenericException("Error al convertir la fecha: " + e.getMessage());
	}
    }

}
