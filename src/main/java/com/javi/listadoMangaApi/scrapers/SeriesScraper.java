package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.SeriesDto;
import com.javi.listadoMangaApi.dto.VolumeDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class SeriesScraper {

    public SeriesDto scrapSeriesPage(int id) throws GenericException, IOException {
	String link = UrlConstants.BASE_URL + UrlConstants.SERIES_PATH + "?" + UrlConstants.ID_PARAM + "=" + id;
	SeriesDto seriesDto = null;

	Document doc = Jsoup.connect(link).get();
	Elements seriesInfo = doc.getElementsByClass("izq");
	String seriesName = seriesInfo.select("h2").get(0).text();
	List<String> dataCollector = new ArrayList<>();
	List<Integer> idCollector = new ArrayList<>();
	List<String> linkCollector = new ArrayList<>();
	Elements seriesData = doc.select("b");
	Elements linkData = seriesInfo.select("a:not([href^='http'])");
	for (int i = 1; i < seriesData.size(); i++) {
	    if (!seriesData.get(i).nextSibling().toString().isEmpty()
		    && !seriesData.get(i).nextSibling().toString().equals(" ")) {
		dataCollector.add(seriesData.get(i).nextSibling().toString().trim());
	    }
	}
	linkData.forEach(data -> {
	    idCollector.add(CommonUtils.getIdFromLink(data));
	    linkCollector.add(data.text());
	});
	List<VolumeDto> volumes = new ArrayList<>();
	Elements volumesInfo = doc.getElementsByClass("cen");

	volumesInfo.forEach(volume -> {
	    String volumeName = "";
	    String volumePages = null;
	    String volumePrice = null;
	    String volumeDate = null;
	    String[] volumeSplit = volume.toString().split("<br>");
	    if (volumeSplit.length > 2) {
		for (int i = volumeSplit.length - 1; i > -1; i--) {
		    if (i == volumeSplit.length - 1) {
			volumeDate = volumeSplit[i].trim();
		    } else if (i == volumeSplit.length - 2) {
			volumePrice = volumeSplit[i].trim();
		    } else if (i == volumeSplit.length - 3) {
			if (volumeSplit[i].contains("páginas")) {
			    volumePages = volumeSplit[i].trim();
			} else {

			    volumeName = volumeName + volumeSplit[i];
			}
		    } else {
			int divLoc = volumeSplit[i].indexOf("/div>");
			volumeName = volumeSplit[i].substring(divLoc + 4, volumeSplit[i].length()) + " "
				+ volumeName.trim();
		    }
		}
	    } else {
		for (String volumeString : volumeSplit) {
		    int divLoc = volumeString.indexOf("/div>");
		    int htmlLoc = volumeString.indexOf("</td>");
		    if (htmlLoc != -1) {
			volumeName = volumeName + " " + volumeString.substring(divLoc + 4, htmlLoc);
		    } else {
			volumeName = volumeName + " " + volumeString.substring(divLoc + 4, volumeString.length());
		    }
		}
	    }
	    if (volumeDate != null) {
		Document volumeDateDoc = Jsoup.parse(volumeDate);
		volumeDate = volumeDateDoc.text();
	    }

	    VolumeDto volumeFinal = new VolumeDto(
		    volumeName.trim().replace("\\n", "").replace(">", "").replace("</a", ""), volumePages, volumePrice,
		    volumeDate);
	    volumes.add(volumeFinal);
	});

	if (dataCollector.size() == 5) {
	    if (idCollector.size() == 5) {
		seriesDto = new SeriesDto(id, seriesName, dataCollector.get(0), null, idCollector.get(0),
			linkCollector.get(0), idCollector.get(1), linkCollector.get(1), 0, null, idCollector.get(2),
			linkCollector.get(2), idCollector.get(3), linkCollector.get(3), idCollector.get(4),
			linkCollector.get(4), dataCollector.get(1), dataCollector.get(2), dataCollector.get(3),
			dataCollector.get(4), volumes);
	    } else {
		seriesDto = new SeriesDto(id, seriesName, dataCollector.get(0), null, idCollector.get(0),
			linkCollector.get(0), idCollector.get(1), linkCollector.get(1), idCollector.get(2),
			linkCollector.get(2), idCollector.get(3), linkCollector.get(3), idCollector.get(4),
			linkCollector.get(4), idCollector.get(5), linkCollector.get(5), dataCollector.get(1),
			dataCollector.get(2), dataCollector.get(3), dataCollector.get(4), volumes);
	    }

	} else {
	    if (idCollector.size() == 5) {
		seriesDto = new SeriesDto(id, seriesName, dataCollector.get(0), dataCollector.get(1),
			idCollector.get(0), linkCollector.get(0), idCollector.get(1), linkCollector.get(1), 0, null,
			idCollector.get(2), linkCollector.get(2), idCollector.get(3), linkCollector.get(3),
			idCollector.get(4), linkCollector.get(4), dataCollector.get(2), dataCollector.get(3),
			dataCollector.get(4), dataCollector.get(5), volumes);
	    } else {
		seriesDto = new SeriesDto(id, seriesName, dataCollector.get(0), dataCollector.get(1),
			idCollector.get(0), linkCollector.get(0), idCollector.get(1), linkCollector.get(1),
			idCollector.get(2), linkCollector.get(2), idCollector.get(3), linkCollector.get(3),
			idCollector.get(4), linkCollector.get(4), idCollector.get(5), linkCollector.get(5),
			dataCollector.get(2), dataCollector.get(3), dataCollector.get(4), dataCollector.get(5),
			volumes);
	    }

	}
	return seriesDto;
    }

}
