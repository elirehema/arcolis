package cellstyles;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ECellStyle {

    public CellStyle getCellStyle(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();

        Font font = cell.getSheet().getWorkbook().createFont();
        font.setFontName("Operator Mono Book");
        font.setColor(IndexedColors.GREEN.index);
        cellStyle.setFont(font);

        /** Set Bottom Styles **/
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

    public Font getCellFont(Sheet sheet){
        /**Set Fonts **/
        Font font = sheet.getWorkbook().createFont();
        font.setBoldweight((short) 12);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());

        return font;
    }
}
