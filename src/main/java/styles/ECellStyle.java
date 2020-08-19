package styles;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

public class ECellStyle {
    private Cell cell;
    private Sheet sheet;


    public ECellStyle(Cell cell) {
        this.cell = cell;
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
        return cellStyle;
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


    public static Map<EType, CellStyle> createStyles(Workbook wb){
        Map<EType, CellStyle> styles = new HashMap<>();
        EFont eFont = new EFont();
        CellStyle style;

        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)18);
        titleFont.setBold(true);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short)11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());

        Font headerFont = wb.createFont();
        headerFont.setFontName("Operator Mono Medium");
        headerFont.getItalic();
        headerFont.setFontHeightInPoints((short) 13);


        /**Default Header, Title and Cell **/
        /**@param HEADER **/

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(titleFont);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        styles.put(EType.DEFAULT_TITLE, style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        style.setFont(headerFont);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styles.put(EType.DEFAULT_HEADER, style);

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
        styles.put(EType.DEFAULT_CELL, style);


        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put(EType.GREY_HEADER, style);

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
        styles.put(EType.DASHED_LIGHT_GREY, style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(EType.FORMULA_1, style);

        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put(EType.FORMULA_2, style);


        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put(EType.SOLID_LIGHT_TORQUES, style);

        /** Dash Dot cell Border style **/
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        style.setBorderBottom(BorderStyle.DASH_DOT_DOT);
        style.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put(EType.DASHED_LIGHT_GREY,style);

        return styles;
    }

}
