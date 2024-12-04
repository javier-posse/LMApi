package com.javi.listadoMangaApi.commons;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.javi.listadoMangaApi.dto.SeriesReleaseDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelGenerator {
    public static String generateExcelReleases(int year, List<SeriesReleaseDto> seriesReleases)
	    throws GenericException {

	Workbook excel = new XSSFWorkbook();
	Sheet sheet = excel.createSheet("Persons");
	sheet.setColumnWidth(0, 8000);
	sheet.setColumnWidth(1, 8000);

	// Headers de la tabla
	Row header = sheet.createRow(0);

	CellStyle headerStyle = excel.createCellStyle();
	headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
	headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	XSSFFont font = ((XSSFWorkbook) excel).createFont();
	font.setFontName("Arial");
	font.setFontHeightInPoints((short) 16);
	font.setBold(true);
	headerStyle.setFont(font);

	Cell headerCell = header.createCell(0);
	headerCell.setCellValue("Nombre Novedad");
	headerCell.setCellStyle(headerStyle);

	headerCell = header.createCell(1);
	headerCell.setCellValue("Fecha Novedad");
	headerCell.setCellStyle(headerStyle);

	// lineas de series
	CellStyle style = excel.createCellStyle();
	style.setWrapText(true);

	for (int i = 0; i < seriesReleases.size(); i++) {

	    Row row = sheet.createRow(i + 1);

	    Cell cell = row.createCell(0);
	    cell.setCellValue(seriesReleases.get(i).getName());
	    cell.setCellStyle(style);

	    cell = row.createCell(1);
	    cell.setCellValue(seriesReleases.get(i).getReleaseDate());
	    cell.setCellStyle(style);
	}
	// cierre y guardado del fichero
	File currDir = new File("./src/main/resources/Generated Excel");
	String path = currDir.getAbsolutePath();
	String fileLocation = path + File.separator + year + "_releases.xlsx";

	FileOutputStream outputStream;
	try {
	    outputStream = new FileOutputStream(fileLocation);
	    excel.write(outputStream);
	    excel.close();
	} catch (IOException e) {
	    throw new GenericException("Ha habido un error no controlado al generar el excel");
	}

	return fileLocation;
    }
}
