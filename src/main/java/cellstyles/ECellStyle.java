package cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.HashMap;
import java.util.Map;

public class ECellStyle {
    private Cell cell;
    private Sheet sheet;

    /**
     * Empty constructor
     **/
    public ECellStyle() {
    }

    public ECellStyle(Sheet sheet) {
        this.sheet = sheet;
    }

    public ECellStyle(Cell cell) {
        this.cell = cell;
    }

    public ECellStyle(Sheet sheet, Cell cell) {
        this.cell = cell;
        this.sheet = sheet;
    }

    public ECellStyle withCell(Cell cell) {
        this.cell = cell;
        return this;
    }

    public ECellStyle withCellAndSheet(Cell cell, Sheet sheet) {
        this.cell = cell;
        this.sheet = sheet;
        return this;
    }

    /**
     * Set Bottom Styles
     **/
    public CellStyle getCellStyle() {
        CellStyle cellStyle = this.cell.getCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT_DOT);
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        this.cell.setCellStyle(cellStyle);
        return cellStyle;
    }

    public CellStyle getCellStyles(Cell cell) {
        this.cell = cell;
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setFont(setTextFonts(cellStyle));
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return cellStyle;
    }

    public ECellStyle setDefaultCellStyle() {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT_DOT);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        cell.setCellStyle(cellStyle);
        return this;
    }


    /**
     * Set Excel sheet default BackgroundStyle.
     **/
    public ECellStyle setDefaultBackgroundStyle() {
        CellStyle backgroundStyle = getSheetCellStyle();
        backgroundStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.cell.setCellStyle(backgroundStyle);
        return this;
    }


    public ECellStyle setDefaultHeaderBackground() {
        CellStyle cellStyle = getSheetCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        this.cell.setCellStyle(cellStyle);
        setHeaderFonts(cellStyle);
        return this;
    }


    /**
     * Apply Header fonts
     **/
    private Font setHeaderFonts(CellStyle cellStyle) {
        Font font = getSheetTextFont();
        font.setBold(true);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());
        cellStyle.setFont(font);
        return font;
    }

    /**
     * Apply fonts
     **/
    private Font setTextFonts(CellStyle cellStyle) {
        Font font = getSheetTextFont();
        font.setFontName("Operator Mono Medium");
        font.setColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFont(font);
        return font;
    }

    /**
     * Set sheet text font styles
     **/
    private Font getSheetTextFont() {
        Font font = this.sheet.getWorkbook().createFont();
        if (font.equals(null)) {
            font = this.cell.getSheet().getWorkbook().createFont();
        }
        return font;
    }

    /**
     * Get CellStyle from Sheet or Cell
     **/
    private CellStyle getSheetCellStyle() {
        CellStyle cellStyle = this.sheet.getWorkbook().createCellStyle();
        if (cellStyle == null) {
            cellStyle = this.cell.getCellStyle();
        }
        return cellStyle;
    }

    public static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short)11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put("cell", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula_2", style);

        return styles;
    }

}
