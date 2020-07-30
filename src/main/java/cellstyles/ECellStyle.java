package cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ECellStyle {
    private Cell cell;
    private Sheet sheet;

    /**
     * Empty constructor
     **/
    public ECellStyle() {
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

    /** Set Bottom Styles **/
    public CellStyle getCellStyle(Cell cell) {
        this.cell = cell;
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setFont(getSheetTextFont());
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        return cellStyle;

    }

    public ECellStyle setDefaultCellStyle() {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setLeftBorderColor(IndexedColors.WHITE.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.WHITE.getIndex());
        return this;
    }

    public ECellStyle setCellStyles() {
        this.cell = cell;
        cell.setCellStyle(getCellStyle(cell));
        return this;
    }

    /**
     * Set Excel sheet default BackgroundStyle.
     **/
    public ECellStyle setDefaultBackgroundStyle() {
        CellStyle backgroundStyle = getCellStyle(cell);
        backgroundStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        backgroundStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell.setCellStyle(backgroundStyle);
        return this;
    }

    /**
     * Set Custom Sheet fonts
     **/
    public ECellStyle setCustomSheetFont(Sheet sheet) {
        setHeaderFonts();
        return this;
    }

    /**
     * Set default sheet font
     **/
    public ECellStyle setDefaultSheetFont() {
        setHeaderFonts();
        return this;
    }

    public ECellStyle setDefaultHeaderBackground() {
        CellStyle cellStyle = cell.getCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        return this;
    }


    /**
     * Apply fonts
     **/
    private ECellStyle setHeaderFonts() {
        Font font = getSheetTextFont();
        font.setBoldweight((short) 12);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());

        return this;
    }

    /**
     * Set sheet text font styles
     **/
    private Font getSheetTextFont() {
        Font font = null;
        if (sheet != null && cell == null) {
            font = sheet.getWorkbook().createFont();

        } else if (cell != null && sheet == null) {
            font = cell.getSheet().getWorkbook().createFont();
        } else {
            System.out.print("Sheet not found");
        }
        cell.getSheet().getWorkbook().createFont();
        font.setFontName("Operator Mono Book");
        font.setColor(IndexedColors.GREEN.index);
        return font;
    }

}
