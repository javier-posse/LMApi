package com.javi.listadoMangaApi.scrapers;

import java.io.IOException;
import java.util.List;

import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.commons.CommonUtils;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.AuthorDto;
import com.javi.listadoMangaApi.dto.SeriesSimplifiedDto;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class AuthorScraper {

    public AuthorDto scrapAuthorPage(int id) throws IOException {
	log.info("AuthorScraper inicializado");
	String link = UrlConstants.BASE_URL + UrlConstants.AUTHOR_PATH + "?" + UrlConstants.ID_PARAM + "=" + id;
	AuthorDto author = null;

	// get name and bio
	Document doc = Jsoup.connect(link).get();
	Elements authorInfo = doc.getElementsByClass("izq");
	String authorName = authorInfo.select("h2").text();
	String bio = authorInfo.after("h2").text();
	// strip the name from the bio
	bio = bio.length() > authorName.length() ? bio.substring(authorName.length() + 1, bio.length()) : "";

	// get series
	List<SeriesSimplifiedDto> series = CommonUtils.getSeriesList(doc);

	author = new AuthorDto(id, authorName, bio, series);

	return author;
    }

}
