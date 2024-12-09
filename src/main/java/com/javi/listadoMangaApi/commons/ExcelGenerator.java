package com.javi.listadoMangaApi.commons;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.javi.listadoMangaApi.dto.SeriesReleaseDto;
import com.javi.listadoMangaApi.exception.GenericException;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelGenerator {
    public static String generateExcelReleases(int year, List<SeriesReleaseDto> seriesReleases)
	    throws GenericException {

	Workbook excel = new XSSFWorkbook();
	Sheet sheet = excel.createSheet("Lanzamientos");
	sheet.setColumnWidth(0, 17000);
	sheet.setColumnWidth(1, 10000);
	sheet.setColumnWidth(2, 10000);
	sheet.setColumnWidth(3, 10000);

	// Headers de la tabla y sus estilo
	Row header = sheet.createRow(0);

	CellStyle headerStyle = excel.createCellStyle();
	headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
	headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	XSSFFont font = ((XSSFWorkbook) excel).createFont();
	font.setFontName("Arial");
	font.setFontHeightInPoints((short) 16);
	font.setBold(true);
	headerStyle.setFont(font);
	headerStyle.setAlignment(HorizontalAlignment.CENTER);
	headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	headerStyle.setBorderBottom(BorderStyle.THICK);
	headerStyle.setBorderTop(BorderStyle.THICK);
	headerStyle.setBorderLeft(BorderStyle.THICK);
	headerStyle.setBorderRight(BorderStyle.THICK);

	Cell headerCell = header.createCell(0);
	headerCell.setCellValue("Nombre");
	headerCell.setCellStyle(headerStyle);

	headerCell = header.createCell(1);
	headerCell.setCellValue("Fecha");
	headerCell.setCellStyle(headerStyle);

	headerCell = header.createCell(2);
	headerCell.setCellValue("Editorial");
	headerCell.setCellStyle(headerStyle);

	headerCell = header.createCell(3);
	headerCell.setCellValue("Tipo");
	headerCell.setCellStyle(headerStyle);

	// lineas de series
	CellStyle style = excel.createCellStyle();
	style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	style.setWrapText(true);
	style.setAlignment(HorizontalAlignment.CENTER);
	style.setVerticalAlignment(VerticalAlignment.CENTER);
	style.setBorderBottom(BorderStyle.THIN);

	// Validaciones para crear el drop-down
	String[] options = { "Tomo único", "Primer tomo", "Último tomo" };

	DataValidationHelper validationHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);
	DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
	CellRangeAddressList addressList = new CellRangeAddressList(1, seriesReleases.size(), 3, 3);
	DataValidation validation = validationHelper.createValidation(constraint, addressList);

	sheet.addValidationData(validation);

	for (int i = 0; i < seriesReleases.size(); i++) {

	    Row row = sheet.createRow(i + 1);

	    Cell cell = row.createCell(0);
	    cell.setCellValue(seriesReleases.get(i).getName());
	    cell.setCellStyle(style);

	    cell = row.createCell(1);
	    cell.setCellValue(seriesReleases.get(i).getReleaseDate());
	    cell.setCellStyle(style);

	    cell = row.createCell(2);
	    cell.setCellValue(seriesReleases.get(i).getPublisherName());
	    cell.setCellStyle(style);

	    cell = row.createCell(3);
	    if (seriesReleases.get(i).isOnlyVolume()) {
		cell.setCellValue(options[0]);
	    } else if (seriesReleases.get(i).isFirstRelease()) {
		cell.setCellValue(options[1]);
	    } else if (seriesReleases.get(i).isLastVolume()) {
		cell.setCellValue(options[2]);
	    }
	    cell.setCellStyle(style);
	}

	// crear tabla para bordes
	CellRangeAddress regionBodyLeft = new CellRangeAddress(1, seriesReleases.size(), 0, 0);
	CellRangeAddress regionBodyCenterLeft = new CellRangeAddress(1, seriesReleases.size(), 1, 1);
	CellRangeAddress regionBodyCenterRight = new CellRangeAddress(1, seriesReleases.size(), 2, 2);
	CellRangeAddress regionBodyRight = new CellRangeAddress(1, seriesReleases.size(), 2, 2);

	// borders body
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyLeft, sheet);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyLeft, sheet);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyLeft, sheet);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyLeft, sheet);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyCenterLeft, sheet);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyCenterLeft, sheet);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyCenterLeft, sheet);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyCenterLeft, sheet);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyCenterRight, sheet);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyCenterRight, sheet);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyCenterRight, sheet);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyCenterRight, sheet);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyRight, sheet);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyRight, sheet);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyRight, sheet);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyRight, sheet);

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
	    throw new GenericException(
		    "Ha habido un error no controlado al generar el excel: " + e.getLocalizedMessage());
	}

	return fileLocation;
    }
}
