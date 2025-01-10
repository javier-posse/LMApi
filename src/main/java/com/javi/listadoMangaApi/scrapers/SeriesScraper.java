package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.AuthorDto;
import com.javi.listadoMangaApi.dto.SeriesDto;
import com.javi.listadoMangaApi.dto.SeriesInfoDto;
import com.javi.listadoMangaApi.dto.VolumeDto;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class SeriesScraper {

    public SeriesDto scrapSeriesPage(int id) throws IOException {
	log.info("SeriesScraper inicializado");
	String link = UrlConstants.BASE_URL + UrlConstants.SERIES_PATH + "?" + UrlConstants.ID_PARAM + "=" + id;
	SeriesDto seriesDto = null;

	Document doc = Jsoup.connect(link).get();
	Elements seriesInfo = doc.getElementsByClass("izq");
	String seriesName = seriesInfo.select("h2").get(0).text();
	Elements volumesInfo = doc.getElementsByClass("cen");
	List<VolumeDto> volumes = getVolumes(volumesInfo);

	SeriesInfoDto info = getSeriesInfo(doc);
	seriesDto = new SeriesDto(id, seriesName, info, volumes);
	return seriesDto;
    }

    private List<VolumeDto> getVolumes(Elements volumesInfo) {
	List<VolumeDto> volumes = new ArrayList<>();
	volumesInfo.forEach(volume -> {
	    VolumeDto volumeFinal = getVolumeInfo(volume);
	    volumes.add(volumeFinal);
	});
	return volumes;
    }

    private VolumeDto getVolumeInfo(Element volume) {
	StringBuilder volumeName = new StringBuilder();
	String volumePages = null;
	String volumePrice = null;
	String volumeDate = null;
	String[] volumeSplit = volume.toString().split("<br>");
	if (volumeSplit.length > 3) {
	    volumeDate = getVolumeDate(volumeSplit);
	    volumePrice = getVolumePrice(volumeSplit);
	    volumePages = getVolumePages(volumeSplit);
	    volumeName.append(getVolumeName(volumeSplit));
	} else {
	    for (String volumeString : volumeSplit) {
		int divLoc = volumeString.indexOf("/div>");
		int htmlLoc = volumeString.indexOf("</td>");
		if (htmlLoc != -1) {
		    volumeName.append(" " + volumeString.substring(divLoc + 4, htmlLoc));
		} else {
		    volumeName.append(" " + volumeString.substring(divLoc + 4, volumeString.length()));
		}
	    }
	}

	return new VolumeDto(volumeName.toString().trim().replace("\\n", "").replace(">", "").replace("</a", ""),
		volumePages, volumePrice, volumeDate);
    }

    private String getVolumeName(String[] volumeSplit) {
	StringBuilder volumeName = new StringBuilder();
	for (int i = 0; i < volumeSplit.length - 2; i++) {
	    if (!volumeSplit[i].contains("páginas")) {
		int divLoc = volumeSplit[i].indexOf("/div>");
		volumeName.append(volumeSplit[i].substring(divLoc + 4, volumeSplit[i].length()) + " ");
	    }
	}
	return volumeName.toString();
    }

    private String getVolumeDate(String[] volumeSplit) {
	String volumeDate = volumeSplit[volumeSplit.length - 1].trim();
	Document volumeDateDoc = Jsoup.parse(volumeDate);
	return volumeDateDoc.text();
    }

    private String getVolumePrice(String[] volumeSplit) {
	return volumeSplit[volumeSplit.length - 2].trim();
    }

    private String getVolumePages(String[] volumeSplit) {
	if (volumeSplit[volumeSplit.length - 3].contains("páginas")) {
	    return volumeSplit[volumeSplit.length - 3].trim();
	}
	return null;
    }

    private SeriesInfoDto getSeriesInfo(Document doc) {
	Element info = doc.selectFirst("td.izq");
	SeriesInfoDto seriesInfo = new SeriesInfoDto();

	if (doc.selectFirst("td.izq") != null) {
	    for (Element element : info.children()) {
		if (element.tagName().equals("b")) {
		    String key = element.text().replace(":", "").trim();
		    Node value = element.nextSibling();

		    if (value.toString().trim().isEmpty() && element.nextSibling() != null
			    && element.nextSibling().nextSibling() != null) {
			value = element.nextSibling().nextSibling();
		    }
		    switch (key) {
		    case "Título original":
			seriesInfo.setOriginalName(value.toString().trim());
			break;
		    case "Guion", "Historia":
			seriesInfo.setScripterName(((Element) value).text());
			seriesInfo.setScripterId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Dibujo", "Ilustraciones":
			seriesInfo.setArtistName(((Element) value).text());
			seriesInfo.setArtistId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Diseño de personajes":
			seriesInfo.setCharacterDesignName(((Element) value).text());
			seriesInfo.setCharacterDesignId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Historia original":
			seriesInfo.setOriginalStoryName(((Element) value).text());
			seriesInfo.setOriginalStoryId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Autores":
			seriesInfo.setAuthors(generateAuthorsList(value));
			break;
		    case "Color":
			seriesInfo.setColorName(((Element) value).text());
			seriesInfo.setColorId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Editorial japonesa", "Editorial francesa", "Editorial americana", "Editorial surcoreana":
			seriesInfo.setJpPublisherName(((Element) value).text());
			seriesInfo.setJpPublisherId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Editorial española":
			seriesInfo.setSpPublisherName(((Element) value).text());
			seriesInfo.setSpPublisherId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Colección":
			seriesInfo.setCollectionName(((Element) value).text());
			seriesInfo.setCollectionId(CommonUtils.getIdFromLink(((Element) value)));
			break;
		    case "Formato":
			seriesInfo.setFormFactor(value.toString().trim());
			break;
		    case "Sentido de lectura":
			seriesInfo.setReedOrder(value.toString().trim());
			break;
		    case "Números en japonés", "Números en francés", "Números en inglés", "Números en coreano",
			    "Números en chino":
			seriesInfo.setJpPublishedVolumes(value.toString().trim());
			break;
		    case "Números en castellano":
			seriesInfo.setSpPublishedVolumes(value.toString().trim());
			break;
		    default:
			log.info("Información no controlada: {}", value);
			break;
		    }
		}
	    }
	}
	return seriesInfo;
    }

    private AuthorDto[] generateAuthorsList(Node author) {
	List<AuthorDto> authors = new ArrayList<>();

	if (((Element) author).tagName().equals("a")) {
	    int authorId = CommonUtils.getIdFromLink(((Element) author));
	    String authorName = ((Element) author).text();
	    authors.add(new AuthorDto(authorId, authorName, null, null));
	}
	Node sibling = author.nextSibling();
	while (sibling != null) {
	    if (sibling instanceof Element element) {
		if (element.tagName().equals("a")) {
		    int authorId = CommonUtils.getIdFromLink((element));
		    String authorName = element.text();
		    authors.add(new AuthorDto(authorId, authorName, null, null));
		} else if (element.tagName().equals("br")) {
		    break;
		}
	    }
	    sibling = sibling.nextSibling();
	}

	return authors.toArray(new AuthorDto[0]);
    }

}
