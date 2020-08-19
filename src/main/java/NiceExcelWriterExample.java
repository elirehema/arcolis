
import styles.ECellStyle;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import styles.EType;
import workbook.EWorkBook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;
import java.util.logging.Logger;

/**
 * -- This file created by eli on 15/07/2020 for poixss
 * --
 * -- Licensed to the Apache Software Foundation (ASF) under one
 * -- or more contributor license agreements. See the NOTICE file
 * -- distributed with this work for additional information
 * -- regarding copyright ownership. The ASF licenses this file
 * -- to you under the Apache License, Version 2.0 (the
 * -- "License"); you may not use this file except in compliance
 * -- with the License. You may obtain a copy of the License at
 * --
 * -- http://www.apache.org/licenses/LICENSE-2.0
 * --
 * -- Unless required by applicable law or agreed to in writing,
 * -- software distributed under the License is distributed on an
 * -- "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * -- KIND, either express or implied. See the License for the
 * -- specific language governing permissions and limitations
 * -- under the License.
 * --
 **/
public class NiceExcelWriterExample {
    private Logger logger = Logger.getLogger(NiceExcelWriterExample.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();
    private Workbook workbook;
    private final Sheet sheet = null;
    Map<EType, CellStyle> styles = null;

    public NiceExcelWriterExample() {
    }

    public static class LazyHolder {
        public static final NiceExcelWriterExample INSTANCE = new NiceExcelWriterExample();
    }

    public NiceExcelWriterExample getInstance() {
        return LazyHolder.INSTANCE;
    }

    public void ExcelSheet(List<?> objectList, String excelFilePath) throws IOException {
        workbook = eWorkBook.getDefaultExcelWorkbook(excelFilePath);
        styles = eCellStyle.createStyles(workbook);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        populateSingleSheetWithData(sheet, objectList);
        try (FileOutputStream outputStream = new FileOutputStream(appendFileExtensionFormatIfNotProvided(excelFilePath))) {
            workbook.write(outputStream);
            outputStream.close();
        }

    }

    @SuppressWarnings("unchecked")
    public void ExcelSheets(List<Map<String, List<?>>> objs, String excelFilePath) throws IOException {
        workbook = eWorkBook.getDefaultExcelWorkbook(excelFilePath);
        styles = eCellStyle.createStyles(workbook);

        for (Map<String, List<?>> objectMap : objs) {
            Iterator iterator = objectMap.entrySet().stream().sorted(Map.Entry.comparingByKey()).iterator();
            while (iterator.hasNext()) {
                Map.Entry entry = (Map.Entry) iterator.next();
                Sheet sheet = workbook.createSheet(entry.getKey().toString());
                List<Object> objects = (List<Object>) entry.getValue();
                populateSingleSheetWithData(sheet, objects);
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream(appendFileExtensionFormatIfNotProvided(excelFilePath))) {
            workbook.write(outputStream);
            outputStream.close();
        }

    }

    private void createHeaderRow(Object o, Sheet sheet, EType headerStyles) {
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
        titleCell.setCellValue(sheet.getSheetName());
        titleCell.setCellStyle(styles.get(EType.TITLE));
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, o.getClass().getDeclaredFields().length));

        Row headerRow = sheet.createRow(1);
        headerRow.setHeightInPoints(35);
        indexcells = headerRow.createCell(0);
        indexcells.setCellStyle(styles.get(headerStyles));
        indexcells.setCellValue("#");
        int index = 0;
        Field[] fields = o.getClass().getDeclaredFields();
        for (int i = 0; i < getFieldNames(fields).size(); i++) {
            dataCell = headerRow.createCell(++index);
            sheet.setColumnWidth(i, (fields[i].getName().length() + 12) * 256);
            dataCell.setCellStyle(styles.get(headerStyles));
            dataCell.setCellValue(fields[i].getName().toUpperCase());
        }
        index = 0;

    }

    /**
     * Write data to created excel sheet
     * Create initial column named #
     * Give {@param rowNumber} to initial column
     **/

    private void writeExcelSheetBook(Object obj, Integer rowNumber, Row row) {
        Cell cell = row.createCell(0);
        eCellStyle = new ECellStyle(cell);
        cell.setCellStyle(eCellStyle.getCellStyle());
        cell.setCellValue(rowNumber.toString());
        cell.setCellStyle(styles.get("cell"));
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(styles.get("cell"));
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

    private void populateSingleSheetWithData(Sheet sheet, List<?> objects) {
        int rowCount = 0;
        for (Object o : objects) {
            createHeaderRow(o, sheet, EType.HEADER);
            Row row = sheet.createRow(++rowCount);
            try {
                writeExcelSheetBook(o, rowCount, row);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        rowCount = 0;
    }


}
