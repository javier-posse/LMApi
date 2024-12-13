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
import com.javi.listadoMangaApi.dto.MonthReleasesDto;
import com.javi.listadoMangaApi.dto.SeriesReleaseDto;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class MonthReleasesScraper {

    public MonthReleasesDto scrapMonthReleasesPage(int month, int year) throws IOException {
	String link = UrlConstants.BASE_URL + UrlConstants.CALENDAR_PATH + "?" + UrlConstants.MONTH_CALENDAR_PARAM + "="
		+ month + "&" + UrlConstants.YEAR_CALENDAR_PARAM + "=" + year;
	MonthReleasesDto monthReleases = null;

	// get name
	Document doc = Jsoup.connect(link).get();
	Elements monthInfo = doc.getElementsByClass("cen");
	String monthName = monthInfo.first().text();

	// get series released
	List<SeriesReleaseDto> seriesReleases = new ArrayList<>();
	Elements tables = doc.select("table");
	Elements releases = tables.select(".izq:not([style])");
	releases.forEach(release -> {
	    Elements children = release.children();
	    for (int i = 0; i < children.size(); i++) {
		int releaseId = 0;
		String releaseName = null;
		int releaseAuthId = 0;
		String releaseAuthName = null;
		int releaseArtistId = 0;
		String releaseArtistName = null;
		boolean lastVolume = false;
		boolean firstRelease = false;
		boolean onlyVolume = false;
		String date = null;
		String publisher = null;
		int cont = 0;
		if (children.get(i).tagName().equals("br")) {
		    while (children.get(i + cont).nextElementSibling() != null
			    && children.get(i + cont).nextElementSibling().tagName().equals("a")) {
			cont++;
		    }
		    if (cont > 0) {
			for (int c = 1; c < cont + 1; c++) {
			    if (c == 1) {
				releaseId = CommonUtils.getLinkId(children.get(i + c));
				releaseName = children.get(i + c).text();
			    } else if (c == 2) {
				releaseAuthId = CommonUtils.getLinkId(children.get(i + c));
				releaseAuthName = children.get(i + c).text();
			    } else if (c == 3) {
				releaseArtistId = CommonUtils.getLinkId(children.get(i + c));
				releaseArtistName = children.get(i + c).text();
			    }
			}
			// recorro para pillar la fecha y cuando llego al elemento paro. Es lo
			// unico que se me ha ocurrido
			Elements allElements = doc.getAllElements();
			for (Element element : allElements) {
			    if (element.attr("style").equals("vertical-align: middle;")) {
				publisher = element.getElementsByTag("h2").get(0).text();
				date = element.getElementsByTag("h2").get(1).text();
			    } else if (element.equals(children.get(i))) {
				break;
			    }
			}
			Elements conditionals = children.get(i + cont + 1).select("span");
			if (!conditionals.isEmpty()) {
			    switch (conditionals.first().text()) {
			    case "ÚLTIMO NÚMERO":
				lastVolume = true;
				break;
			    case "NOVEDAD":
				firstRelease = true;
				break;
			    case "NÚMERO ÚNICO":
				onlyVolume = true;
				break;
			    default:
				break;
			    }
			}
			SeriesReleaseDto seriesReleaseDto = new SeriesReleaseDto(releaseId, releaseName, date,
				publisher, releaseAuthId, releaseAuthName, releaseArtistId, releaseArtistName,
				lastVolume, firstRelease, onlyVolume);
			seriesReleases.add(seriesReleaseDto);
		    }
		}
	    }

	});
	monthReleases = new MonthReleasesDto(monthName, seriesReleases);

	return monthReleases;
    }
}
