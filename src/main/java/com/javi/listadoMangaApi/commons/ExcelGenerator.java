package com.javi.listadoMangaApi.commons;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Instant;
import java.time.LocalDate;
import java.time.Month;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.javi.listadoMangaApi.config.ExcelConfig;
import com.javi.listadoMangaApi.constants.ExcelConstants;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.SeriesReleaseDto;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@Component
public class ExcelGenerator {
    private ExcelGenerator() {
    }

    public String generateExcelReleases(int year, List<SeriesReleaseDto> seriesReleases, int lastVolumeCount)
	    throws IOException {

	log.info("ExcelGenerator inicializado");

	//Excel File
	Path path = Paths.get(ExcelConstants.FILEPATH);
	Files.createDirectories(path);
	DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern(ExcelConstants.DATETIME_FORMAT_FILE_NAME);
	String fileLocation = path.toString() + File.separator + year + ExcelConstants.RELEASES
		+ Instant.now().atZone(ZoneId.of(ExcelConstants.MADRID_DATETIME)).format(dateFormat) + ExcelConstants.FILE_EXTENSION;

	XSSFWorkbook excel = new XSSFWorkbook();
	int[] conts = { 0, 0, 0 };
	CellStyle headerStyle = createHeaderStyle(excel);
	CellStyle tableBodyStyle = createTableBodyStyle(excel);
	generateFirstSheet(excel, conts, seriesReleases, headerStyle, tableBodyStyle);
	generateSecondSheet(excel, conts, seriesReleases, headerStyle, tableBodyStyle, lastVolumeCount, year);
	generateThirdSheet(excel, conts);

	FileOutputStream outputStream;
	outputStream = new FileOutputStream(fileLocation);
	excel.write(outputStream);
	excel.close();
	outputStream.close();

	return fileLocation;
    }

    private XSSFSheet generateFirstSheet(XSSFWorkbook excel, int[] contMangaMania,
	    List<SeriesReleaseDto> seriesReleases, CellStyle headerStyle, CellStyle tableBodyStyle) {
	XSSFSheet sheetReleases = excel.createSheet(ExcelConstants.FIRST_SHEET_NAME);
	sheetReleases.setColumnWidth(0, 17000);
	sheetReleases.setColumnWidth(1, 10000);
	sheetReleases.setColumnWidth(2, 10000);
	sheetReleases.setColumnWidth(3, 10000);

	// Headers de la tabla y sus estilo
	Row header = sheetReleases.createRow(0);
	Cell headerCell = header.createCell(0);
	headerCell.setCellValue(ExcelConstants.TABLE_COLUMN_NAME_LABEL);
	headerCell.setCellStyle(headerStyle);
	headerCell = header.createCell(1);
	headerCell.setCellValue(ExcelConstants.TABLE_COLUMN_DATE_LABEL);
	headerCell.setCellStyle(headerStyle);
	headerCell = header.createCell(2);
	headerCell.setCellValue(ExcelConstants.TABLE_COLUMN_PUBLISHER_LABEL);
	headerCell.setCellStyle(headerStyle);
	headerCell = header.createCell(3);
	headerCell.setCellValue(ExcelConstants.TABLE_COLUMN_TYPE_LABEL);
	headerCell.setCellStyle(headerStyle);

	CellStyle cellStyleDate = createDatesStyle(excel);

	String[] options = { ExcelConstants.TABLE_TYPE_DROPDOWN_ONLY_VOLUME, ExcelConstants.TABLE_TYPE_DROPDOWN_FIRST_VOLUME, ExcelConstants.TABLE_TYPE_DROPDOWN_LAST_VOLUME };

	DataValidationHelper validationHelper = new XSSFDataValidationHelper(sheetReleases);
	DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
	CellRangeAddressList addressList = new CellRangeAddressList(1, seriesReleases.size(), ExcelConstants.Cell.THIRD, ExcelConstants.Cell.THIRD);
	DataValidation validation = validationHelper.createValidation(constraint, addressList);

	sheetReleases.addValidationData(validation);

	for (int i = 0; i < seriesReleases.size(); i++) {

	    Row row = sheetReleases.createRow(i + 1);

	    Cell cell = row.createCell(ExcelConstants.Cell.ZERO);
	    cell.setCellValue(seriesReleases.get(i).getName());
	    Hyperlink link = excel.getCreationHelper().createHyperlink(HyperlinkType.URL);
	    link.setAddress(CommonUtils.getLinkFromId(seriesReleases.get(i).getId(), UrlConstants.SERIES_PATH));
	    cell.setHyperlink(link);

	    if (seriesReleases.get(i).getName().contains(ExcelConstants.MANIA)
		    && seriesReleases.get(i).getPublisherName().equals(ExcelConstants.PLANETA)) {
		contMangaMania[0]++;
	    }
	    cell.setCellStyle(tableBodyStyle);

	    cell = row.createCell(ExcelConstants.Cell.FIRST);
	    // Convierto la fecha a formato fecha de Java
	    String releaseDate = seriesReleases.get(i).getReleaseDate();
	    cell.setCellValue(CommonUtils.fullDateConverter(releaseDate));
	    cell.setCellStyle(cellStyleDate);

	    cell = row.createCell(ExcelConstants.Cell.SECOND);
	    cell.setCellValue(seriesReleases.get(i).getPublisherName());
	    cell.setCellStyle(tableBodyStyle);

	    cell = row.createCell(ExcelConstants.Cell.THIRD);
	    if (seriesReleases.get(i).isOnlyVolume()) {
		cell.setCellValue(options[0]);
	    } else if (seriesReleases.get(i).isFirstRelease()) {
		cell.setCellValue(options[1]);
	    } else if (seriesReleases.get(i).isLastVolume()) {
		cell.setCellValue(options[2]);
	    }
	    cell.setCellStyle(tableBodyStyle);
	}

	sheetReleases.setAutoFilter(new CellRangeAddress(0, seriesReleases.size(), ExcelConstants.Cell.FIRST, ExcelConstants.Cell.FIRST));

	return sheetReleases;
    }

    private XSSFSheet generateSecondSheet(XSSFWorkbook excel, int[] conts, List<SeriesReleaseDto> seriesReleases,
	    CellStyle headerStyle, CellStyle tableBodyStyle, int lastVolumeCount, int year) {
	XSSFSheet sheetReleases = excel.getSheetAt(0);
	// Generado hoja de estadisticas
	XSSFSheet sheetStatistics = excel.createSheet(ExcelConstants.SECOND_SHEET_NAME);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.ZERO, 15000);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.FIRST, 10000);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.SECOND, 10000);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.THIRD, 10000);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.FOURTH, 10000);
	sheetStatistics.setColumnWidth(ExcelConstants.Cell.FIFTH, 10000);

	// tabla de editoriales
	// Contar las apariciones de cada editorial en el primer sheet
	Map<String, Integer> publisherCounts = new HashMap<>();
	for (Row rowPublishers : sheetReleases) {
	    Cell cell = rowPublishers.getCell(ExcelConstants.Cell.SECOND);
	    if (cell != null && cell.getCellType() == CellType.STRING) {
		String publisher = cell.getStringCellValue();
		publisherCounts.put(publisher, publisherCounts.getOrDefault(publisher, 0) + 1);
	    }
	}
	// Ordeno las editoriales de más a menos por número de apariciones
	List<Map.Entry<String, Integer>> sortedEditorialCounts = new ArrayList<>(publisherCounts.entrySet());
	sortedEditorialCounts.sort((entry1, entry2) -> entry2.getValue().compareTo(entry1.getValue()));
	// Colocar en el sheet las editoriales
	Row headerRow = sheetStatistics.createRow(conts[1]++);
	headerRow.createCell(ExcelConstants.Cell.ZERO).setCellValue(ExcelConstants.TABLE_COLUMN_PUBLISHER_NAME_LABEL);
	headerRow.createCell(ExcelConstants.Cell.FIRST).setCellValue(ExcelConstants.TABLE_COLUMN_NEW_RELEASES_QUANTITY_LABEL);
	headerRow.getCell(ExcelConstants.Cell.ZERO).setCellStyle(headerStyle);
	headerRow.getCell(ExcelConstants.Cell.FIRST).setCellStyle(headerStyle);

	for (Map.Entry<String, Integer> entry : sortedEditorialCounts) {
	    Row rowPublishersTable = sheetStatistics.createRow(conts[1]++);
	    rowPublishersTable.createCell(ExcelConstants.Cell.ZERO).setCellValue(entry.getKey());
	    rowPublishersTable.createCell(ExcelConstants.Cell.FIRST).setCellValue(entry.getValue());
	    rowPublishersTable.getCell(ExcelConstants.Cell.ZERO).setCellStyle(tableBodyStyle);
	    rowPublishersTable.getCell(ExcelConstants.Cell.FIRST).setCellStyle(tableBodyStyle);
	}
	// total de tomos unicos y primeros tomos
	int onlyVolumeCont = 0;
	int firstVolumeCont = 0;
	for (SeriesReleaseDto serie : seriesReleases) {
	    if (serie.isOnlyVolume()) {
		onlyVolumeCont++;
	    }
	    if (serie.isFirstRelease()) {
		firstVolumeCont++;
	    }
	}
	// total de lanzamientos
	Cell cellReleasesLabel = sheetStatistics.getRow(1).createCell(ExcelConstants.Cell.THIRD);
	cellReleasesLabel.setCellValue(ExcelConstants.CELL_TOTAL_RELEASES_LABEL);
	cellReleasesLabel.setCellStyle(headerStyle);
	Cell cellReleasesCount = sheetStatistics.getRow(1).createCell(ExcelConstants.Cell.FOURTH);
	cellReleasesCount.setCellValue(seriesReleases.size());

	// Total de tomos únicos
	Cell cellOnlyVolumeLabel = sheetStatistics.getRow(4).createCell(ExcelConstants.Cell.THIRD);
	cellOnlyVolumeLabel.setCellValue(ExcelConstants.CELL_TOTAL_ONLY_VOLUMES_LABEL);
	cellOnlyVolumeLabel.setCellStyle(headerStyle);
	Cell cellOnlyVolumeCount = sheetStatistics.getRow(4).createCell(ExcelConstants.Cell.FOURTH);
	cellOnlyVolumeCount.setCellValue(onlyVolumeCont);

	// Total de primeros tomos
	Cell cellFirstVolumeLabel = sheetStatistics.getRow(7).createCell(ExcelConstants.Cell.THIRD);
	cellFirstVolumeLabel.setCellValue(ExcelConstants.CELL_TOTAL_NEW_SERIES_LABEL);
	cellFirstVolumeLabel.setCellStyle(headerStyle);
	Cell cellFirstVolumeCount = sheetStatistics.getRow(7).createCell(ExcelConstants.Cell.FOURTH);
	cellFirstVolumeCount.setCellValue(firstVolumeCont);

	// Total de terminadas
	Cell cellLastVolumeLabel = sheetStatistics.getRow(10).createCell(ExcelConstants.Cell.THIRD);
	cellLastVolumeLabel.setCellValue(ExcelConstants.CELL_TOTAL_LAST_VOLUMES_LABEL);
	cellLastVolumeLabel.setCellStyle(headerStyle);
	Cell cellLastVolumeCount = sheetStatistics.getRow(10).createCell(ExcelConstants.Cell.FOURTH);
	cellLastVolumeCount.setCellValue(lastVolumeCount);

	// Total de manga manía
	Cell cellMangaManiaLabel = sheetStatistics.getRow(13).createCell(ExcelConstants.Cell.THIRD);
	cellMangaManiaLabel.setCellValue(ExcelConstants.CELL_TOTAL_MANGA_MANIA_LABEL);
	cellMangaManiaLabel.setCellStyle(headerStyle);
	Cell celllMangaManiaCount = sheetStatistics.getRow(13).createCell(ExcelConstants.Cell.FOURTH);
	celllMangaManiaCount.setCellValue(conts[0]);

	// Total de editoriales
	Cell cellEditorialesLabel = sheetStatistics.getRow(16).createCell(ExcelConstants.Cell.THIRD);
	cellEditorialesLabel.setCellValue(ExcelConstants.CELL_TOTAL_PUBLISHERS_LABEL);
	cellEditorialesLabel.setCellStyle(headerStyle);
	Cell celllEditorialesCount = sheetStatistics.getRow(16).createCell(ExcelConstants.Cell.FOURTH);
	celllEditorialesCount.setCellValue(sortedEditorialCounts.size());

	createMonthTable(conts, sheetReleases, sheetStatistics, headerStyle, tableBodyStyle);
	generatePastYearsReleasesTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearsClosedTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearsOnlyVolumesTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearNewSeriesTable(sheetStatistics, year, headerStyle, tableBodyStyle);

	return sheetStatistics;

    }

    private void generatePastYearsReleasesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(19);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(19);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(ExcelConstants.Cell.THIRD);
	cellReleasesHeaderYear.setCellValue(ExcelConstants.TABLE_COLUMN_YEAR_LABEL);
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(ExcelConstants.Cell.FOURTH);
	cellReleasesHeaderQtty.setCellValue(ExcelConstants.TABLE_COLUMN_NEWS_LABEL);
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_RELEASES + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(ExcelConstants.Cell.THIRD);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(ExcelConstants.Cell.FOURTH);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_RELEASES + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private void generatePastYearsClosedTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(ExcelConstants.Cell.CLOSED_SERIES);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(ExcelConstants.Cell.CLOSED_SERIES);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(ExcelConstants.Cell.THIRD);
	cellReleasesHeaderYear.setCellValue(ExcelConstants.TABLE_COLUMN_YEAR_LABEL);
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(ExcelConstants.Cell.FOURTH);
	cellReleasesHeaderQtty.setCellValue(ExcelConstants.TABLE_COLUMN_CLOSED_LABEL);
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_CLOSED + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(ExcelConstants.Cell.THIRD);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(ExcelConstants.Cell.FOURTH);
		cellReleasesQtty.setCellValue(Integer.valueOf(config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_CLOSED + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private void generatePastYearsOnlyVolumesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(ExcelConstants.Cell.ONLY_VOLUMES);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(ExcelConstants.Cell.ONLY_VOLUMES);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(ExcelConstants.Cell.THIRD);
	cellReleasesHeaderYear.setCellValue(ExcelConstants.TABLE_COLUMN_YEAR_LABEL);
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(ExcelConstants.Cell.FOURTH);
	cellReleasesHeaderQtty.setCellValue(ExcelConstants.TABLE_COLUMN_ONLY_VOLUMES_LABEL);
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_ONLY_VOLUMES + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(ExcelConstants.Cell.THIRD);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(ExcelConstants.Cell.FOURTH);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_ONLY_VOLUMES + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private void generatePastYearNewSeriesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(ExcelConstants.Cell.NEW_SERIES);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(ExcelConstants.Cell.NEW_SERIES);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(ExcelConstants.Cell.THIRD);
	cellReleasesHeaderYear.setCellValue(ExcelConstants.TABLE_COLUMN_YEAR_LABEL);
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(ExcelConstants.Cell.FOURTH);
	cellReleasesHeaderQtty.setCellValue(ExcelConstants.TABLE_COLUMN_NEW_RELEASES_LABEL);
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_NEW_SERIES + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(ExcelConstants.Cell.THIRD);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(ExcelConstants.Cell.FOURTH);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty(ExcelConstants.EXCEL_PROPERTY_YEAR_NEW_SERIES + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private XSSFSheet generateThirdSheet(XSSFWorkbook excel, int[] conts) {
	XSSFSheet sheetStatistics = excel.getSheetAt(1);
	XSSFSheet sheetCharts = excel.createSheet(ExcelConstants.THIRD_SHEET_NAME);
	createCharts(sheetStatistics, sheetCharts, conts);
	return sheetCharts;
    }

    private void createCharts(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	createPublisherChart(sheetStatistics, sheetCharts, conts);
	createMonthChart(sheetStatistics, sheetCharts, conts);
	createYearReleasesChart(sheetStatistics, sheetCharts);
	createYearClosedChart(sheetStatistics, sheetCharts);
	createYearOnlyVolumesChart(sheetStatistics, sheetCharts);
	createYearNewSeriesChart(sheetStatistics, sheetCharts);
    }

    private void createPublisherChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	// Generación de gráfico editoriales
	XSSFDrawing pieChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart pieChart = pieChartDrawing.createChart(pieChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 8, ExcelConstants.Row.PUBLISHER_CHART));
	pieChart.setTitleText(ExcelConstants.GRAPH_PUBLISHER_LABEL);
	pieChart.setTitleOverlay(false);
	pieChart.getOrAddLegend().setPosition(LegendPosition.RIGHT);

	XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(1, conts[1] > 10 ? 10 : conts[1] - 1, 0, 0));
	XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(1, conts[1] > 10 ? 10 : conts[1] - 1, 1, 1));

	XDDFPieChartData data = (XDDFPieChartData) pieChart.createData(ChartTypes.PIE, null, null);
	data.addSeries(categories, values).setTitle(ExcelConstants.GRAPH_PUBLISHER_LEYEND_LABEL, null);
	pieChart.plot(data);
    }

    private void createMonthTable(int[] conts, XSSFSheet sheetReleases, XSSFSheet sheetStatistics,
	    CellStyle headerStyle, CellStyle tableBodyStyle) {
	// genero un array de meses para crear un hashmap y así tenerlos ordenaicos

	Locale localeEs = Locale.of("es", "ES");
	Map<String, Integer> novedadesPorMes = new LinkedHashMap<>();
        for (Month mes : Month.values()) {
            String nombreMes = mes.getDisplayName(TextStyle.FULL, localeEs).toLowerCase(localeEs);
            novedadesPorMes.put(nombreMes, 0);
        }
	// Contaje de repeticiones de mes
	DateTimeFormatter monthFormatter = DateTimeFormatter.ofPattern("MMMM", Locale.forLanguageTag("es-ES"));
	for (Row rowi : sheetReleases) {
	    Cell cell = rowi.getCell(ExcelConstants.Cell.FIRST);
	    if (cell != null && cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
		LocalDate date = cell.getLocalDateTimeCellValue().toLocalDate();
		String month = date.format(monthFormatter).toLowerCase();
		novedadesPorMes.put(month, novedadesPorMes.get(month) + 1);
	    }
	}

	conts[2] = conts[1] + 1;
	Row tempRow = sheetStatistics.createRow(conts[2]++);
	tempRow.createCell(0).setCellValue(ExcelConstants.TABLE_COLUMN_MONTH_LABEL);
	tempRow.createCell(1).setCellValue(ExcelConstants.TABLE_COLUMN_NEWS_LABEL);
	tempRow.getCell(0).setCellStyle(headerStyle);
	tempRow.getCell(1).setCellStyle(headerStyle);
	for (Map.Entry<String, Integer> entry : novedadesPorMes.entrySet()) {
	    tempRow = sheetStatistics.createRow(conts[2]++);
	    tempRow.createCell(0).setCellValue(StringUtils.capitalize(entry.getKey()));
	    tempRow.createCell(1).setCellValue(entry.getValue());
	    tempRow.getCell(0).setCellStyle(tableBodyStyle);
	    tempRow.getCell(1).setCellStyle(tableBodyStyle);
	}
    }

    private void createMonthChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {

	XSSFDrawing monthChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart monthChart = monthChartDrawing.createChart(monthChartDrawing.createAnchor(0, 0, 0, 0, 0, ExcelConstants.Row.MONTH_CHART_START, 15, ExcelConstants.Row.MONTH_CHART_END));
	monthChart.setTitleText(ExcelConstants.GRAPH_BY_MONTH_LABEL);
	monthChart.setTitleOverlay(false);
	monthChart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = monthChart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = monthChart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> months = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(conts[1] + 2, conts[2] - 1, 0, 0));
	XDDFNumericalDataSource<Double> releases = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(conts[1] + 2, conts[2] - 1, 1, 1));

	XDDFBarChartData mothData = (XDDFBarChartData) monthChart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
	mothData.addSeries(months, releases).setTitle(ExcelConstants.GRAPH_BY_MONTH_LEYEND_LABEL, null);
	monthChart.plot(mothData);
    }

    private void createYearReleasesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, ExcelConstants.Row.YEAR_RELEASES_CHART_START, 15, ExcelConstants.Row.YEAR_RELEASES_CHART_END));

	chart.plot(
		generateChartData(sheetStatistics, ExcelConstants.GRAPH_LAST_YEAR_RELEASES_LABEL, ExcelConstants.GRAPH_LAST_YEAR_RELEASES_LEYEND_LABEL, chart, 20));
    }

    private void createYearClosedChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, ExcelConstants.Row.YEAR_CLOSED_CHART_START, 15, ExcelConstants.Row.YEAR_CLOSED_CHART_END));

	chart.plot(generateChartData(sheetStatistics, ExcelConstants.GRAPH_LAST_YEAR_CLOSED_LABEL, ExcelConstants.GRAPH_LAST_YEAR_CLOSED_LEYEND_LABEL,
		chart, 30));
    }

    private void createYearOnlyVolumesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, ExcelConstants.Row.YEAR_ONLY_VOLUMES_CHART_START, 15, ExcelConstants.Row.YEAR_ONLY_VOLUMES_CHART_END));

	chart.plot(generateChartData(sheetStatistics, ExcelConstants.GRAPH_LAST_YEAR_ONLY_VOLUMES_LABEL, ExcelConstants.GRAPH_LAST_YEAR_ONLY_VOLUMES_LEYEND_LABEL, chart,
		40));
    }

    private void createYearNewSeriesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, ExcelConstants.Row.YEAR_NEW_SERIES_CHART_START, 15, ExcelConstants.Row.YEAR_NEW_SERIES_CHART_END));

	chart.plot(generateChartData(sheetStatistics, ExcelConstants.GRAPH_LAST_YEAR_NEW_LABEL, ExcelConstants.GRAPH_LAST_YEAR_NEW_LEYEND_LABEL,
		chart, 50));
    }

    private XDDFChartData generateChartData(XSSFSheet sheetStatistics, String chartName, String chartTitle,
	    XSSFChart chart, int dataInitColum) {

	chart.setTitleText(chartName);
	chart.setTitleOverlay(false);
	chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> dataX = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(dataInitColum, dataInitColum + 4, ExcelConstants.Cell.THIRD, ExcelConstants.Cell.THIRD));
	XDDFNumericalDataSource<Double> dataY = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(dataInitColum, dataInitColum + 4, ExcelConstants.Cell.FOURTH, ExcelConstants.Cell.FOURTH));
	XDDFChartData chartData = chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
	chartData.addSeries(dataX, dataY).setTitle(chartTitle, null);

	return chartData;
    }

    private CellStyle createHeaderStyle(XSSFWorkbook excel) {
	CellStyle headerStyle = excel.createCellStyle();
	headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
	headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	XSSFFont font = (excel).createFont();
	font.setFontName(ExcelConstants.FONT);
	font.setFontHeightInPoints((short) 16);
	font.setBold(true);
	headerStyle.setFont(font);
	headerStyle.setAlignment(HorizontalAlignment.CENTER);
	headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
	headerStyle.setBorderBottom(BorderStyle.THICK);
	headerStyle.setBorderTop(BorderStyle.THICK);
	headerStyle.setBorderLeft(BorderStyle.THICK);
	headerStyle.setBorderRight(BorderStyle.THICK);

	return headerStyle;
    }

    private CellStyle createTableBodyStyle(XSSFWorkbook excel) {
	// lineas de series
	CellStyle style = excel.createCellStyle();
	style.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	style.setWrapText(true);
	style.setAlignment(HorizontalAlignment.CENTER);
	style.setVerticalAlignment(VerticalAlignment.CENTER);

	style.setBorderBottom(BorderStyle.THIN);
	style.setBorderTop(BorderStyle.THIN);
	style.setBorderLeft(BorderStyle.THIN);
	style.setBorderRight(BorderStyle.THIN);

	return style;
    }

    private CellStyle createDatesStyle(XSSFWorkbook excel) {
	CellStyle cellStyleDate = excel.createCellStyle();
	CreationHelper createHelper = excel.getCreationHelper();
	cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat(ExcelConstants.DATETIME_FORMAT));
	cellStyleDate.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
	cellStyleDate.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cellStyleDate.setWrapText(true);
	cellStyleDate.setAlignment(HorizontalAlignment.CENTER);
	cellStyleDate.setVerticalAlignment(VerticalAlignment.CENTER);
	cellStyleDate.setBorderBottom(BorderStyle.THIN);
	return cellStyleDate;
    }
}
