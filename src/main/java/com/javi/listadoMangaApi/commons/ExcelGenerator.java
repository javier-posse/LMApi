package com.javi.listadoMangaApi.commons;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
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
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.javi.listadoMangaApi.dto.SeriesReleaseDto;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class ExcelGenerator {
    // constructor privado para que no se genere el público default
    private ExcelGenerator() {
    }

    public static String generateExcelReleases(int year, List<SeriesReleaseDto> seriesReleases, int lastVolumeCount)
	    throws IOException {

	// Genero lo necesario para guardar el fichero
	File currDir = new File("./");
	String path = currDir.getAbsolutePath();
	DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss");
	String fileLocation = path + File.separator + year + "_releases_"
		+ Instant.now().atZone(ZoneId.of("Europe/Madrid")).format(dateFormat) + ".xlsx";

	XSSFWorkbook excel = new XSSFWorkbook();
	int[] conts = { 0, 0, 0 };
	CellStyle headerStyle = createHeaderStyle(excel);
	XSSFSheet sheetReleases = generateFirstSheet(excel, conts, seriesReleases, headerStyle);
	XSSFSheet sheetStatistics = generateSecondSheet(excel, conts, seriesReleases, headerStyle, sheetReleases,
		lastVolumeCount);
	generateThirdSheet(excel, sheetStatistics, conts);

	// cierre y guardado del fichero
	FileOutputStream outputStream;
	outputStream = new FileOutputStream(fileLocation);
	excel.write(outputStream);
	excel.close();
	outputStream.close();

	return fileLocation;
    }

    private static XSSFSheet generateFirstSheet(XSSFWorkbook excel, int[] contMangaMania,
	    List<SeriesReleaseDto> seriesReleases, CellStyle headerStyle) {
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

	// lineas de series
	CellStyle style = excel.createCellStyle();
	style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	style.setWrapText(true);
	style.setAlignment(HorizontalAlignment.CENTER);
	style.setVerticalAlignment(VerticalAlignment.CENTER);
	style.setBorderBottom(BorderStyle.THIN);

	// lineas de fechas
	CellStyle cellStyleDate = excel.createCellStyle();
	CreationHelper createHelper = excel.getCreationHelper();
	cellStyleDate.setDataFormat(createHelper.createDataFormat().getFormat("dd MMMM yyyy"));
	cellStyleDate.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
	cellStyleDate.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	cellStyleDate.setWrapText(true);
	cellStyleDate.setAlignment(HorizontalAlignment.CENTER);
	cellStyleDate.setVerticalAlignment(VerticalAlignment.CENTER);
	cellStyleDate.setBorderBottom(BorderStyle.THIN);

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
	    if (seriesReleases.get(i).getName().contains("Manía")
		    && seriesReleases.get(i).getPublisherName().equals("Planeta Cómic")) {
		contMangaMania[0]++;
	    }
	    cell.setCellStyle(style);

	    cell = row.createCell(1);
	    // Convierto la fecha a formato fecha de Java
	    String releaseDate = seriesReleases.get(i).getReleaseDate();
	    cell.setCellValue(CommonUtils.fullDateConverter(releaseDate));
	    cell.setCellStyle(cellStyleDate);

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
	CellRangeAddress regionBodyRight = new CellRangeAddress(1, seriesReleases.size(), 3, 3);

	// borders body
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyLeft, sheetReleases);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyLeft, sheetReleases);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyLeft, sheetReleases);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyLeft, sheetReleases);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyCenterLeft, sheetReleases);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyCenterLeft, sheetReleases);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyCenterLeft, sheetReleases);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyCenterLeft, sheetReleases);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyCenterRight, sheetReleases);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyCenterRight, sheetReleases);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyCenterRight, sheetReleases);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyCenterRight, sheetReleases);
	RegionUtil.setBorderTop(BorderStyle.THIN, regionBodyRight, sheetReleases);
	RegionUtil.setBorderBottom(BorderStyle.THIN, regionBodyRight, sheetReleases);
	RegionUtil.setBorderLeft(BorderStyle.THIN, regionBodyRight, sheetReleases);
	RegionUtil.setBorderRight(BorderStyle.THIN, regionBodyRight, sheetReleases);

	return sheetReleases;
    }

    private static XSSFSheet generateSecondSheet(XSSFWorkbook excel, int[] conts, List<SeriesReleaseDto> seriesReleases,
	    CellStyle headerStyle, XSSFSheet sheetReleases, int lastVolumeCount) {
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

	for (Map.Entry<String, Integer> entry : sortedEditorialCounts) {
	    Row rowPublishersTable = sheetStatistics.createRow(conts[1]++);
	    rowPublishersTable.createCell(0).setCellValue(entry.getKey());
	    rowPublishersTable.createCell(1).setCellValue(entry.getValue());
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
	Row row = sheetStatistics.getRow(1);
	Cell cellReleasesLabel = row.createCell(3);
	cellReleasesLabel.setCellValue("Total lanzamientos: ");
	cellReleasesLabel.setCellStyle(headerStyle);
	Cell cellReleasesCount = row.createCell(4);
	cellReleasesCount.setCellValue(seriesReleases.size());

	// Total de tomos únicos
	Row row4 = sheetStatistics.getRow(4);
	Cell cellOnlyVolumeLabel = row4.createCell(3);
	cellOnlyVolumeLabel.setCellValue("Total tomos únicos: ");
	cellOnlyVolumeLabel.setCellStyle(headerStyle);
	Cell cellOnlyVolumeCount = row4.createCell(4);
	cellOnlyVolumeCount.setCellValue(onlyVolumeCont);

	// Total de primeros tomos
	Row row7 = sheetStatistics.getRow(7);
	Cell cellFirstVolumeLabel = row7.createCell(3);
	cellFirstVolumeLabel.setCellValue("Total nuevas series: ");
	cellFirstVolumeLabel.setCellStyle(headerStyle);
	Cell cellFirstVolumeCount = row7.createCell(4);
	cellFirstVolumeCount.setCellValue(firstVolumeCont);

	// Total de terminadas
	Row row10 = sheetStatistics.getRow(10);
	Cell cellLastVolumeLabel = row10.createCell(3);
	cellLastVolumeLabel.setCellValue("Total series terminadas: ");
	cellLastVolumeLabel.setCellStyle(headerStyle);
	Cell cellLastVolumeCount = row10.createCell(4);
	cellLastVolumeCount.setCellValue(lastVolumeCount);

	// Total de manga manía
	Row row13 = sheetStatistics.getRow(13);
	Cell cellMangaManiaLabel = row13.createCell(3);
	cellMangaManiaLabel.setCellValue("Total \"Manga Manía\": ");
	cellMangaManiaLabel.setCellStyle(headerStyle);
	Cell celllMangaManiaCount = row13.createCell(4);
	celllMangaManiaCount.setCellValue(conts[0]);

	createMonthTable(conts, sheetReleases, sheetStatistics);

	return sheetStatistics;

    }

    private static XSSFSheet generateThirdSheet(XSSFWorkbook excel, XSSFSheet sheetStatistics, int[] conts) {
	XSSFSheet sheetCharts = excel.createSheet("Gráficas");
	createCharts(sheetStatistics, sheetCharts, conts);
	return sheetCharts;
    }

    private static void createCharts(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	createPublisherChart(sheetStatistics, sheetCharts, conts);
	createMonthChart(sheetStatistics, sheetCharts, conts);

    }

    private static void createPublisherChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {
	// Generación de gráfico editoriales
	XSSFDrawing pieChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFClientAnchor anchorPie = pieChartDrawing.createAnchor(0, 0, 0, 0, 0, 2, 8, 14);
	XSSFChart pieChart = pieChartDrawing.createChart(anchorPie);
	pieChart.setTitleText("Distribución de Editoriales");
	pieChart.setTitleOverlay(false);
	XDDFChartLegend legend = pieChart.getOrAddLegend();
	legend.setPosition(LegendPosition.RIGHT);

	XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(4, conts[1] - 1, 0, 0));
	XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(4, conts[1] - 1, 1, 1));
	XDDFPieChartData data = (XDDFPieChartData) pieChart.createData(ChartTypes.PIE, null, null);
	XDDFPieChartData.Series series = (XDDFPieChartData.Series) data.addSeries(categories, values);
	series.setTitle("Editoriales", null);
	pieChart.plot(data);
    }

    private static void createMonthTable(int[] conts, XSSFSheet sheetReleases, XSSFSheet sheetStatistics) {
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
	for (Map.Entry<String, Integer> entry : novedadesPorMes.entrySet()) {
	    tempRow = sheetStatistics.createRow(conts[2]++);
	    tempRow.createCell(0).setCellValue(StringUtils.capitalize(entry.getKey()));
	    tempRow.createCell(1).setCellValue(entry.getValue());
	}
    }

    private static void createMonthChart(XSSFSheet sheetStatistics, XSSFSheet sheetCharts, int[] conts) {

	XSSFDrawing monthChartDrawing = sheetCharts.createDrawingPatriarch();
	XSSFClientAnchor monthAnchor = monthChartDrawing.createAnchor(0, 0, 0, 0, 0, 17, 15, 35);
	XSSFChart monthChart = monthChartDrawing.createChart(monthAnchor);
	monthChart.setTitleText("Lanzamientos por meses");
	monthChart.setTitleOverlay(false);
	XDDFChartLegend monthLegend = monthChart.getOrAddLegend();
	monthLegend.setPosition(LegendPosition.BOTTOM);

	XDDFCategoryAxis bottomAxis = monthChart.createCategoryAxis(AxisPosition.LEFT);
	XDDFValueAxis leftAxis = monthChart.createValueAxis(AxisPosition.BOTTOM);
	leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
	int initialIndex = conts[1] + 2;

	XDDFDataSource<String> meses = XDDFDataSourcesFactory.fromStringCellRange(sheetStatistics,
		new CellRangeAddress(initialIndex, conts[2] - 1, 0, 0));
	XDDFNumericalDataSource<Double> novedades = XDDFDataSourcesFactory.fromNumericCellRange(sheetStatistics,
		new CellRangeAddress(initialIndex, conts[2] - 1, 1, 1));

	XDDFBarChartData mothData = (XDDFBarChartData) monthChart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
	XDDFBarChartData.Series monthSeries = (XDDFBarChartData.Series) mothData.addSeries(meses, novedades);
	monthSeries.setTitle("Novedades por Mes", null);
	monthChart.plot(mothData);
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
}
