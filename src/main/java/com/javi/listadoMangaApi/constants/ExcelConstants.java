package com.javi.listadoMangaApi.constants;

public class ExcelConstants {
    public static final String FILEPATH = "./Excel";
    public static final String DATETIME_FORMAT_FILE_NAME = "yyyy-MM-dd%HH-mm-ss";
    public static final String DATETIME_FORMAT = "MMMM/dd/yyyy";
    public static final String FONT = "Arial";
    public static final String RELEASES = "_releases_";
    public static final String MADRID_DATETIME = "Europe/Madrid";
    public static final String FILE_EXTENSION = ".xlsx";
    public static final String FIRST_SHEET_NAME = "Lanzamientos";
    public static final String SECOND_SHEET_NAME = "Estadísticas";
    public static final String THIRD_SHEET_NAME = "Gráficas";
    public static final String TABLE_COLUMN_NAME_LABEL= "Nombre";
    public static final String TABLE_COLUMN_DATE_LABEL= "Fecha";
    public static final String TABLE_COLUMN_PUBLISHER_LABEL= "Editorial";
    public static final String TABLE_COLUMN_TYPE_LABEL= "Tipo";
    public static final String TABLE_TYPE_DROPDOWN_ONLY_VOLUME= "Tomo único";
    public static final String TABLE_TYPE_DROPDOWN_FIRST_VOLUME= "Primer tomo";
    public static final String TABLE_TYPE_DROPDOWN_LAST_VOLUME= "Último tomo";
    public static final String MANIA = "Manía";
    public static final String PLANETA = "Planeta Cómic";
    public static final String TABLE_COLUMN_PUBLISHER_NAME_LABEL = "Nombre editorial";
    public static final String TABLE_COLUMN_NEW_RELEASES_QUANTITY_LABEL = "Cantidad de nuevos lanzamientos";
    public static final String CELL_TOTAL_RELEASES_LABEL = "Total lanzamientos: ";
    public static final String CELL_TOTAL_ONLY_VOLUMES_LABEL = "Total tomos únicos: ";
    public static final String CELL_TOTAL_NEW_SERIES_LABEL = "Total nuevas series: ";
    public static final String CELL_TOTAL_LAST_VOLUMES_LABEL = "Total series terminadas: ";
    public static final String CELL_TOTAL_MANGA_MANIA_LABEL = "Total \"Manga Manía\": ";
    public static final String CELL_TOTAL_PUBLISHERS_LABEL = "Total editoriales: ";
    public static final String TABLE_COLUMN_YEAR_LABEL = "Año";
    public static final String TABLE_COLUMN_NEWS_LABEL = "Novedades";
    public static final String TABLE_COLUMN_CLOSED_LABEL = "Series Cerradas";
    public static final String TABLE_COLUMN_ONLY_VOLUMES_LABEL = "Tomos únicos";
    public static final String TABLE_COLUMN_NEW_RELEASES_LABEL = "Series nuevas";
    public static final String TABLE_COLUMN_MONTH_LABEL = "Mes";
    public static final String EXCEL_PROPERTY_YEAR_RELEASES = "year.releases.";
    public static final String EXCEL_PROPERTY_YEAR_CLOSED = "year.closed.";
    public static final String EXCEL_PROPERTY_YEAR_ONLY_VOLUMES = "year.onlyVolumes.";
    public static final String EXCEL_PROPERTY_YEAR_NEW_SERIES = "year.newSeries.";
    public static final String GRAPH_PUBLISHER_LABEL = "Distribución de Editoriales";
    public static final String GRAPH_PUBLISHER_LEYEND_LABEL = "Editoriales";
    public static final String GRAPH_BY_MONTH_LABEL = "Lanzamientos por meses";
    public static final String GRAPH_BY_MONTH_LEYEND_LABEL = "Novedades por Mes";
    public static final String GRAPH_LAST_YEAR_RELEASES_LABEL = "Lanzamientos años anteriores";
    public static final String GRAPH_LAST_YEAR_RELEASES_LEYEND_LABEL = "Novedades en los años";
    public static final String GRAPH_LAST_YEAR_CLOSED_LABEL = "Series cerradas años anteriores";
    public static final String GRAPH_LAST_YEAR_CLOSED_LEYEND_LABEL = "Series cerradas en los años";
    public static final String GRAPH_LAST_YEAR_ONLY_VOLUMES_LABEL = "Tomos únicos años anteriores";
    public static final String GRAPH_LAST_YEAR_ONLY_VOLUMES_LEYEND_LABEL = "Tomos únicos en los años";
    public static final String GRAPH_LAST_YEAR_NEW_LABEL = "Series nuevas años anteriores";
    public static final String GRAPH_LAST_YEAR_NEW_LEYEND_LABEL = "Series nuevas en los años";


    public static class Cell {
        public static final int ZERO = 0;
        public static final int FIRST = 1;
        public static final int SECOND = 2;
        public static final int THIRD = 3;
        public static final int FOURTH = 4;
        public static final int FIFTH = 5;
        public static final int CLOSED_SERIES = 29;
        public static final int ONLY_VOLUMES = 39;
        public static final int NEW_SERIES = 49;
    }

    public static class Row {
        public static final int FIRST = 1;
        public static final int PUBLISHER_CHART = 14;
        public static final int MONTH_CHART_START = 17;
        public static final int MONTH_CHART_END = 35;
        public static final int YEAR_RELEASES_CHART_START = 38;
        public static final int YEAR_RELEASES_CHART_END = 48;
        public static final int YEAR_CLOSED_CHART_START = 51;
        public static final int YEAR_CLOSED_CHART_END = 61;
        public static final int YEAR_ONLY_VOLUMES_CHART_START = 63;
        public static final int YEAR_ONLY_VOLUMES_CHART_END = 73;
        public static final int YEAR_NEW_SERIES_CHART_START = 76;
        public static final int YEAR_NEW_SERIES_CHART_END = 86;
    }
}
