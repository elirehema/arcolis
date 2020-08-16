package examples;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class AllignCells {
    public static void main(String[] args) throws IOException {
        try {
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet();
            Row row = sheet.createRow(2);
            for (int i = 0; i < 8; i++) {
                sheet.setColumnWidth(1, 256 * 15);
            }

            createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
            createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
            createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
            createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
            createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
            createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
            createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

            // Write the output to a file
            try (OutputStream fileOut = new FileOutputStream("ss-example-align.xlsx")) {
                wb.write(fileOut);
            }
        }catch (IOException e){
            throw e;
        }
    }
    private static void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign) {
        CreationHelper ch = wb.getCreationHelper();
        Cell cell = row.createCell(column);
        cell.setCellValue(ch.createRichTextString("Align It"));
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }
}
