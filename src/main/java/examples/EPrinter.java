package examples;

import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import styles.ECellStyle;
import styles.EType;
import workbook.EWorkBook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class EPrinter {
    private Logger logger = Logger.getLogger(EPrinter.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();
    private Workbook workbook;
    private final Sheet sheet = null;
    Map<EType, CellStyle> styles = null;

    public static class LazyHolder {
        public static final EPrinter INSTANCE = new EPrinter();
    }

    public EPrinter getInstance() {
        return EPrinter.LazyHolder.INSTANCE;
    }

    public void ExcelSheet(
            List<?> objectList, String excelFilePath, EType HeaderCellStyle,
            EType TitleCellStyles, EType DataCellStyle) throws IOException {
        workbook = eWorkBook.getDefaultExcelWorkbook(excelFilePath);
        styles = eCellStyle.createStyles(workbook);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        populateSingleSheetWithData(sheet, objectList, HeaderCellStyle, TitleCellStyles, DataCellStyle);
        try (FileOutputStream outputStream = new FileOutputStream(appendFileExtensionFormatIfNotProvided(excelFilePath))) {
            workbook.write(outputStream);
            outputStream.close();
        }

    }

    @SuppressWarnings("unchecked")
    public void ExcelSheets(
            List<Map<String, List<?>>> objs, String excelFilePath, EType HeaderCellStyle,
            EType TitleCellStyles, EType DataCellStyle) throws IOException {
        workbook = eWorkBook.getDefaultExcelWorkbook(excelFilePath);
        styles = eCellStyle.createStyles(workbook);

        for (Map<String, List<?>> objectMap : objs) {
            Iterator iterator = objectMap.entrySet().stream().sorted(Map.Entry.comparingByKey()).iterator();
            while (iterator.hasNext()) {
                Map.Entry entry = (Map.Entry) iterator.next();
                Sheet sheet = workbook.createSheet(entry.getKey().toString());
                List<Object> objects = (List<Object>) entry.getValue();
                populateSingleSheetWithData(sheet, objects, HeaderCellStyle, TitleCellStyles, DataCellStyle);
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream(appendFileExtensionFormatIfNotProvided(excelFilePath))) {
            workbook.write(outputStream);
            outputStream.close();
        }

    }

    private void createHeaderRow(Object o, Sheet sheet, EType HeaderCellStyle,
                                 EType TitleCellStyles) {
        Cell dataCell, indexcells = null;
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);
        printSetup.setLandscape(true);

        /**Set Page Number on Footer **/
        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        sheet.setDefaultColumnWidth(16);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        sheet.setHorizontallyCenter(true);
        sheet.removeMergedRegion(0);

        Row titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue(sheet.getSheetName().substring(0, 1).toUpperCase() + sheet.getSheetName().substring(1));
        titleCell.setCellStyle(styles.getOrDefault(TitleCellStyles, styles.get(EType.DEFAULT_TITLE)));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, o.getClass().getDeclaredFields().length));

        Row headerRow = sheet.createRow(1);
        headerRow.setHeightInPoints(35);
        indexcells = headerRow.createCell(0);
        indexcells.setCellStyle(styles.getOrDefault(HeaderCellStyle, styles.get(EType.DEFAULT_HEADER)));
        indexcells.setCellValue("#");
        int index = 0;
        Field[] fields = o.getClass().getDeclaredFields();
        for (int i = 0; i < getFieldNames(fields).size(); i++) {
            dataCell = headerRow.createCell(++index);
            sheet.setColumnWidth(i, (fields[i].getName().length() + 12) * 256);
            dataCell.setCellStyle(styles.getOrDefault(HeaderCellStyle, styles.get(EType.DEFAULT_HEADER)));
            dataCell.setCellValue(fields[i].getName().toUpperCase());
        }
        index = 0;

    }

    /**
     * Write data to created excel sheet
     * Create initial column named #
     * Give {@param rowNumber} to initial column
     **/

    private void writeExcelSheetBook(Object obj, Integer rowNumber, Row row, EType DataCellStyle) {
        Cell cell = row.createCell(0);
        eCellStyle = new ECellStyle(cell);
        cell.setCellValue(rowNumber.toString());
        cell.setCellStyle(styles.getOrDefault(DataCellStyle, styles.get(EType.DEFAULT_CELL)));
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(styles.getOrDefault(DataCellStyle, styles.get(EType.DEFAULT_CELL)));
                cell.setCellValue(value.toString());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    private static List<String> getFieldNames(Field[] fields) {
        List<String> fieldNames = new ArrayList<>();
        for (Field field : fields) {
            fieldNames.add(field.getName());
        }

        return fieldNames;
    }


    /**
     * Get default excel format
     **/
    private String appendFileExtensionFormatIfNotProvided(String extension) {
        String filename = null;
        if (extension.endsWith(".xls")) {
            filename = extension;
        } else if (extension.endsWith(".xlsx")) {
            filename = extension;
        } else {
            filename = extension.concat(".xls");
        }

        return filename;
    }

    private void populateSingleSheetWithData(Sheet sheet, List<?> objects, EType HeaderCellStyle, EType TitleCellStyles, EType DataCellStyles) {
        int rowCount = 0;
        for (Object o : objects) {
            createHeaderRow(o, sheet, HeaderCellStyle, TitleCellStyles);
            Row row = sheet.createRow(++rowCount);
            try {
                writeExcelSheetBook(o, rowCount, row, DataCellStyles);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        rowCount = 0;
    }
}
