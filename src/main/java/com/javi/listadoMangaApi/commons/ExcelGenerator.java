package com.javi.listadoMangaApi.commons;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
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
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.javi.listadoMangaApi.config.ExcelConfig;
import com.javi.listadoMangaApi.constants.UrlConstants;
import com.javi.listadoMangaApi.dto.SeriesReleaseDto;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelGenerator {
    // constructor privado para que no se genere el público default
    private ExcelGenerator() {
    }

    public static String generateExcelReleases(int year, List<SeriesReleaseDto> seriesReleases, int lastVolumeCount)
	    throws IOException {

	log.info("ExcelGenerator inicializado");

	// Genero lo necesario para guardar el fichero
	Path path = Paths.get("./Excel");
	Files.createDirectories(path);
	DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("yyyy-MM-dd%HH-mm-ss");
	String fileLocation = path.toString() + File.separator + year + "_releases_"
		+ Instant.now().atZone(ZoneId.of("Europe/Madrid")).format(dateFormat) + ".xlsx";

	XSSFWorkbook excel = new XSSFWorkbook();
	int[] conts = { 0, 0, 0 };
	CellStyle headerStyle = createHeaderStyle(excel);
	CellStyle tableBodyStyle = createTableBodyStyle(excel);
	generateFirstSheet(excel, conts, seriesReleases, headerStyle, tableBodyStyle);
	generateSecondSheet(excel, conts, seriesReleases, headerStyle, tableBodyStyle, lastVolumeCount, year);
	generateThirdSheet(excel, conts);

	// cierre y guardado del fichero
	FileOutputStream outputStream;
	outputStream = new FileOutputStream(fileLocation);
	excel.write(outputStream);
	excel.close();
	outputStream.close();

	return fileLocation;
    }

    private static XSSFSheet generateFirstSheet(XSSFWorkbook excel, int[] contMangaMania,
	    List<SeriesReleaseDto> seriesReleases, CellStyle headerStyle, CellStyle tableBodyStyle) {
	XSSFSheet sheetReleases = excel.createSheet("Lanzamientos");
	sheetReleases.setColumnWidth(0, 17000);
	sheetReleases.setColumnWidth(1, 10000);
	sheetReleases.setColumnWidth(2, 10000);
	sheetReleases.setColumnWidth(3, 10000);

	// Headers de la tabla y sus estilo
	Row header = sheetReleases.createRow(0);

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

	// Estilo de fechas
	CellStyle cellStyleDate = createDatesStyle(excel);

	// Validaciones para crear el drop-down
	String[] options = { "Tomo único", "Primer tomo", "Último tomo" };

	DataValidationHelper validationHelper = new XSSFDataValidationHelper(sheetReleases);
	DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(options);
	CellRangeAddressList addressList = new CellRangeAddressList(1, seriesReleases.size(), 3, 3);
	DataValidation validation = validationHelper.createValidation(constraint, addressList);

	sheetReleases.addValidationData(validation);

	for (int i = 0; i < seriesReleases.size(); i++) {

	    Row row = sheetReleases.createRow(i + 1);

	    Cell cell = row.createCell(0);
	    cell.setCellValue(seriesReleases.get(i).getName());
	    Hyperlink link = excel.getCreationHelper().createHyperlink(HyperlinkType.URL);
	    link.setAddress(CommonUtils.getLinkFromId(seriesReleases.get(i).getId(), UrlConstants.SERIES_PATH));
	    cell.setHyperlink(link);

	    if (seriesReleases.get(i).getName().contains("Manía")
		    && seriesReleases.get(i).getPublisherName().equals("Planeta Cómic")) {
		contMangaMania[0]++;
	    }
	    cell.setCellStyle(tableBodyStyle);

	    cell = row.createCell(1);
	    // Convierto la fecha a formato fecha de Java
	    String releaseDate = seriesReleases.get(i).getReleaseDate();
	    cell.setCellValue(CommonUtils.fullDateConverter(releaseDate));
	    cell.setCellStyle(cellStyleDate);

	    cell = row.createCell(2);
	    cell.setCellValue(seriesReleases.get(i).getPublisherName());
	    cell.setCellStyle(tableBodyStyle);

	    cell = row.createCell(3);
	    if (seriesReleases.get(i).isOnlyVolume()) {
		cell.setCellValue(options[0]);
	    } else if (seriesReleases.get(i).isFirstRelease()) {
		cell.setCellValue(options[1]);
	    } else if (seriesReleases.get(i).isLastVolume()) {
		cell.setCellValue(options[2]);
	    }
	    cell.setCellStyle(tableBodyStyle);
	}

	// filtros de tabla
	sheetReleases.setAutoFilter(new CellRangeAddress(0, seriesReleases.size(), 1, 3));

	return sheetReleases;
    }

    private static XSSFSheet generateSecondSheet(XSSFWorkbook excel, int[] conts, List<SeriesReleaseDto> seriesReleases,
	    CellStyle headerStyle, CellStyle tableBodyStyle, int lastVolumeCount, int year) {
	XSSFSheet sheetReleases = excel.getSheetAt(0);
	// Generado hoja de estadisticas
	XSSFSheet sheetStatistics = excel.createSheet("Estadísticas");
	sheetStatistics.setColumnWidth(0, 15000);
	sheetStatistics.setColumnWidth(1, 10000);
	sheetStatistics.setColumnWidth(2, 10000);
	sheetStatistics.setColumnWidth(3, 10000);
	sheetStatistics.setColumnWidth(4, 10000);
	sheetStatistics.setColumnWidth(5, 10000);

	// tabla de editoriales
	// Contar las apariciones de cada editorial en el primer sheet
	Map<String, Integer> publisherCounts = new HashMap<>();
	for (Row rowPublishers : sheetReleases) {
	    Cell cell = rowPublishers.getCell(2);
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
	headerRow.createCell(0).setCellValue("Nombre editorial");
	headerRow.createCell(1).setCellValue("Cantidad de lanzamientos");
	headerRow.getCell(0).setCellStyle(headerStyle);
	headerRow.getCell(1).setCellStyle(headerStyle);

	for (Map.Entry<String, Integer> entry : sortedEditorialCounts) {
	    Row rowPublishersTable = sheetStatistics.createRow(conts[1]++);
	    rowPublishersTable.createCell(0).setCellValue(entry.getKey());
	    rowPublishersTable.createCell(1).setCellValue(entry.getValue());
	    rowPublishersTable.getCell(0).setCellStyle(tableBodyStyle);
	    rowPublishersTable.getCell(1).setCellStyle(tableBodyStyle);
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
	Cell cellReleasesLabel = sheetStatistics.getRow(1).createCell(3);
	cellReleasesLabel.setCellValue("Total lanzamientos: ");
	cellReleasesLabel.setCellStyle(headerStyle);
	Cell cellReleasesCount = sheetStatistics.getRow(1).createCell(4);
	cellReleasesCount.setCellValue(seriesReleases.size());

	// Total de tomos únicos
	Cell cellOnlyVolumeLabel = sheetStatistics.getRow(4).createCell(3);
	cellOnlyVolumeLabel.setCellValue("Total tomos únicos: ");
	cellOnlyVolumeLabel.setCellStyle(headerStyle);
	Cell cellOnlyVolumeCount = sheetStatistics.getRow(4).createCell(4);
	cellOnlyVolumeCount.setCellValue(onlyVolumeCont);

	// Total de primeros tomos
	Cell cellFirstVolumeLabel = sheetStatistics.getRow(7).createCell(3);
	cellFirstVolumeLabel.setCellValue("Total nuevas series: ");
	cellFirstVolumeLabel.setCellStyle(headerStyle);
	Cell cellFirstVolumeCount = sheetStatistics.getRow(7).createCell(4);
	cellFirstVolumeCount.setCellValue(firstVolumeCont);

	// Total de terminadas
	Cell cellLastVolumeLabel = sheetStatistics.getRow(10).createCell(3);
	cellLastVolumeLabel.setCellValue("Total series terminadas: ");
	cellLastVolumeLabel.setCellStyle(headerStyle);
	Cell cellLastVolumeCount = sheetStatistics.getRow(10).createCell(4);
	cellLastVolumeCount.setCellValue(lastVolumeCount);

	// Total de manga manía
	Cell cellMangaManiaLabel = sheetStatistics.getRow(13).createCell(3);
	cellMangaManiaLabel.setCellValue("Total \"Manga Manía\": ");
	cellMangaManiaLabel.setCellStyle(headerStyle);
	Cell celllMangaManiaCount = sheetStatistics.getRow(13).createCell(4);
	celllMangaManiaCount.setCellValue(conts[0]);

	// Total de editoriales
	Cell cellEditorialesLabel = sheetStatistics.getRow(16).createCell(3);
	cellEditorialesLabel.setCellValue("Total editoriales: ");
	cellEditorialesLabel.setCellStyle(headerStyle);
	Cell celllEditorialesCount = sheetStatistics.getRow(16).createCell(4);
	celllEditorialesCount.setCellValue(sortedEditorialCounts.size());

	createMonthTable(conts, sheetReleases, sheetStatistics, headerStyle, tableBodyStyle);
	generatePastYearsReleasesTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearsClosedTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearsOnlyVolumesTable(sheetStatistics, year, headerStyle, tableBodyStyle);
	generatePastYearNewSeriesTable(sheetStatistics, year, headerStyle, tableBodyStyle);

	return sheetStatistics;

    }

    private static void generatePastYearsReleasesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(19);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(19);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(3);
	cellReleasesHeaderYear.setCellValue("Año");
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(4);
	cellReleasesHeaderQtty.setCellValue("Novedades");
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty("year.releases." + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(3);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(4);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty("year.releases." + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private static void generatePastYearsClosedTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(29);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(29);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(3);
	cellReleasesHeaderYear.setCellValue("Año");
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(4);
	cellReleasesHeaderQtty.setCellValue("Series cerradas");
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty("year.closed." + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(3);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(4);
		cellReleasesQtty.setCellValue(Integer.valueOf(config.getProperty("year.closed." + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private static void generatePastYearsOnlyVolumesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(39);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(39);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(3);
	cellReleasesHeaderYear.setCellValue("Año");
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(4);
	cellReleasesHeaderQtty.setCellValue("Tomos únicos");
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty("year.onlyVolumes." + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(3);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(4);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty("year.onlyVolumes." + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private static void generatePastYearNewSeriesTable(XSSFSheet sheetStatistics, int year, CellStyle headerStyle,
	    CellStyle tableBodyStyle) {
	Row rowHeader = sheetStatistics.getRow(49);
	if (rowHeader == null) {
	    rowHeader = sheetStatistics.createRow(49);
	}
	Cell cellReleasesHeaderYear = rowHeader.createCell(3);
	cellReleasesHeaderYear.setCellValue("Año");
	cellReleasesHeaderYear.setCellStyle(headerStyle);
	Cell cellReleasesHeaderQtty = rowHeader.createCell(4);
	cellReleasesHeaderQtty.setCellValue("Series nuevas");
	cellReleasesHeaderQtty.setCellStyle(headerStyle);

	ExcelConfig config = new ExcelConfig();
	int contYears = 5;
	for (int i = 1; i < 6; i++) {
	    if (config.getProperty("year.newSeries." + (year - contYears)) != null) {
		Row row = sheetStatistics.getRow(rowHeader.getRowNum() + i);
		if (row == null) {
		    row = sheetStatistics.createRow(rowHeader.getRowNum() + i);
		}
		Cell cellReleasesYear = row.createCell(3);
		cellReleasesYear.setCellValue(year - contYears);
		cellReleasesYear.setCellStyle(tableBodyStyle);
		Cell cellReleasesQtty = row.createCell(4);
		cellReleasesQtty
			.setCellValue(Integer.valueOf(config.getProperty("year.newSeries." + (year - contYears))));
		cellReleasesQtty.setCellStyle(tableBodyStyle);
		contYears--;
	    }
	}
    }

    private static XSSFSheet generateThirdSheet(XSSFWorkbook excel, int[] conts) {
	XSSFSheet sheetStatistics = excel.getSheetAt(1);
	XSSFSheet sheetCharts = excel.createSheet("Gráficas");
	createCharts(sheetStatistics, sheetCharts, conts);
	return sheetCharts;
    }

    private static void createCharts(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	createPublisherChart(sheetStatistics, sheetCharts, conts);
	createMonthChart(sheetStatistics, sheetCharts, conts);
	createYearReleasesChart(sheetStatistics, sheetCharts);
	createYearClosedChart(sheetStatistics, sheetCharts);
	createYearOnlyVolumesChart(sheetStatistics, sheetCharts);
	createYearNewSeriesChart(sheetStatistics, sheetCharts);
    }

    private static void createPublisherChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	// Generación de gráfico editoriales
	XSSFDrawing pieChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart pieChart = pieChartDrawing.createChart(pieChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 8, 14));
	pieChart.setTitleText("Distribución de Editoriales");
	pieChart.setTitleOverlay(false);
	pieChart.getOrAddLegend().setPosition(LegendPosition.RIGHT);

	XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(1, conts[1] > 10 ? 10 : conts[1] - 1, 0, 0));
	XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(1, conts[1] > 10 ? 10 : conts[1] - 1, 1, 1));

	XDDFPieChartData data = (XDDFPieChartData) pieChart.createData(ChartTypes.PIE, null, null);
	data.addSeries(categories, values).setTitle("Editoriales", null);
	pieChart.plot(data);
    }

    private static void createMonthTable(int[] conts, XSSFSheet sheetReleases, XSSFSheet sheetStatistics,
	    CellStyle headerStyle, CellStyle tableBodyStyle) {
	// genero un array de meses para crear un hashmap y así tenerlos ordenaicos
	String[] mesesOrdenados = { "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
		"septiembre", "octubre", "noviembre", "diciembre" };
	Map<String, Integer> novedadesPorMes = new LinkedHashMap<>();
	for (String mes : mesesOrdenados) {
	    novedadesPorMes.put(mes, 0);
	}
	// Contaje de repeticiones de mes
	DateTimeFormatter monthFormatter = DateTimeFormatter.ofPattern("MMMM", Locale.forLanguageTag("es-ES"));
	for (Row rowi : sheetReleases) {
	    Cell cell = rowi.getCell(1);
	    if (cell != null && cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
		LocalDate date = cell.getLocalDateTimeCellValue().toLocalDate();
		String month = date.format(monthFormatter);
		novedadesPorMes.put(month, novedadesPorMes.get(month) + 1);
	    }
	}

	conts[2] = conts[1] + 1;
	Row tempRow = sheetStatistics.createRow(conts[2]++);
	tempRow.createCell(0).setCellValue("Mes");
	tempRow.createCell(1).setCellValue("Novedades");
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

    private static void createMonthChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {

	XSSFDrawing monthChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart monthChart = monthChartDrawing.createChart(monthChartDrawing.createAnchor(0, 0, 0, 0, 0, 17, 15, 35));
	monthChart.setTitleText("Lanzamientos por meses");
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
	mothData.addSeries(months, releases).setTitle("Novedades por Mes", null);
	monthChart.plot(mothData);
    }

    private static void createYearReleasesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, 38, 15, 48));
	chart.setTitleText("Lanzamientos años anteriores");
	chart.setTitleOverlay(false);
	chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> years = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(20, 24, 3, 3));
	XDDFNumericalDataSource<Double> releases = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(20, 24, 4, 4));

	XDDFLineChartData chartData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
	chartData.addSeries(years, releases).setTitle("Novedades en los años", null);
	chart.plot(chartData);
    }

    private static void createYearClosedChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, 51, 15, 61));
	chart.setTitleText("Series cerradas años anteriores");
	chart.setTitleOverlay(false);
	chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> years = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(30, 34, 3, 3));
	XDDFNumericalDataSource<Double> closed = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(30, 34, 4, 4));

	XDDFLineChartData chartData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
	chartData.addSeries(years, closed).setTitle("Series cerradas en los años", null);
	chart.plot(chartData);
    }

    private static void createYearOnlyVolumesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, 63, 15, 73));
	chart.setTitleText("Tomos únicos años anteriores");
	chart.setTitleOverlay(false);
	chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> years = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(40, 44, 3, 3));
	XDDFNumericalDataSource<Double> closed = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(40, 44, 4, 4));

	XDDFLineChartData chartData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
	chartData.addSeries(years, closed).setTitle("Tomos únicos en los años", null);
	chart.plot(chartData);
    }

    private static void createYearNewSeriesChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts) {

	XSSFDrawing yearReleasesChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFChart chart = yearReleasesChartDrawing
		.createChart(yearReleasesChartDrawing.createAnchor(0, 0, 0, 0, 0, 76, 15, 86));
	chart.setTitleText("Series nuevas años anteriores");
	chart.setTitleOverlay(false);
	chart.getOrAddLegend().setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);

	XDDFDataSource<String> years = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(50, 54, 3, 3));
	XDDFNumericalDataSource<Double> closed = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(50, 54, 4, 4));

	XDDFLineChartData chartData = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
	chartData.addSeries(years, closed).setTitle("Series nuevas en los años", null);
	chart.plot(chartData);
    }

    private static CellStyle createHeaderStyle(XSSFWorkbook excel) {
	CellStyle headerStyle = excel.createCellStyle();
	headerStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
	headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

	XSSFFont font = (excel).createFont();
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

	return headerStyle;
    }

    private static CellStyle createTableBodyStyle(XSSFWorkbook excel) {
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

    private static CellStyle createDatesStyle(XSSFWorkbook excel) {
	CellStyle cellStyleDate = excel.createCellStyle();
	CreationHelper createHelper = excel.getCreationHelper();
	cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("MMMM/dd/yyyy"));
	cellStyleDate.setFillForegroundColor(IndexedColors.SEA_GREEN.getIndex());
	cellStyleDate.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cellStyleDate.setWrapText(true);
	cellStyleDate.setAlignment(HorizontalAlignment.CENTER);
	cellStyleDate.setVerticalAlignment(VerticalAlignment.CENTER);
	cellStyleDate.setBorderBottom(BorderStyle.THIN);
	return cellStyleDate;
    }
}
