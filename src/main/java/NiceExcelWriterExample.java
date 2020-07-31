import cellstyles.ECellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import workbook.EWorkBook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
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
    private static final Logger LOGGER = Logger.getLogger(NiceExcelWriterExample.class.getName());
    private ECellStyle eCellStyle = null;
    private EWorkBook eWorkBook = new EWorkBook();

    /**
     * Write an Excel File with single Sheet
     **/
    public void writeExcel(List<?> objectList, String excelFilePath) throws IOException {
        Workbook workbook = eWorkBook.getWorkbook(excelFilePath);
        Sheet sheet = workbook.createSheet(excelFilePath.toLowerCase());
        int rowCount = 0;
        for (Object object : objectList) {
            createHeaderRow(object, sheet);
            Row row = sheet.createRow(++rowCount);
            try {
                writeBook(object, rowCount, row);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }
    }

    public void writeMultipleSheetExcel(List<?> languages, String excelFilePath) throws IOException {
        Workbook workbook = eWorkBook.getWorkbook(excelFilePath);
        for (Object parentObject : languages) {
            System.out.println(parentObject.getClass().getDeclaredFields());
            // Sheet sheet = workbook.createSheet("arentObject");
            int rowCount = 0;
            Field[] fields = parentObject.getClass().getDeclaredFields();
            for (Field f : fields) {

                try {
                    Field field = parentObject.getClass().getDeclaredField(f.getName().toString());

                    field.setAccessible(true);
                    Object object = field.get(parentObject);

                    if (object instanceof Collection) {
                        // System.out.println("Collection Instance" + object.toString());
                    }
                    if (object instanceof String) {
                        //  System.out.println("String Instance:  " + object.toString());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
           /* for (Object obj : parentObject.getBooks()) {
                createHeaderRow(obj, sheet);
                Row row = sheet.createRow(++rowCount);
                try {
                    writeBook(obj, rowCount, row);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }*/
            rowCount = 0;
        }
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }
    }

    private void createHeaderRow(Object o, Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Cell cell =  null;
        PrintSetup printSetup = sheet.getPrintSetup();

        /**Set Page Number on Footer **/
        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        /**Set Print Format**/
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);


        Row row = sheet.createRow(0);
        sheet.setDefaultColumnWidth(15);
        sheet.setFitToPage(true);

        cell = row.createCell(0);
        eCellStyle = new ECellStyle(sheet, cell);
        eCellStyle.setDefaultHeaderBackground();
        cell.setCellValue("#");
        int index = 0;
        for (Field f : o.getClass().getDeclaredFields()) {
            cell = row.createCell(++index);
            eCellStyle = new ECellStyle(sheet, cell);
            eCellStyle
                    .setDefaultHeaderBackground();
            cell.setCellValue(f.getName().toUpperCase());
        }

    }

    private void writeBook(Object obj, Integer rowNumber, Row row) {
        Cell cell = row.createCell(0);
        eCellStyle = new ECellStyle(cell);
        cell.setCellValue(rowNumber.toString());
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            eCellStyle = new ECellStyle(cell);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellStyle(eCellStyle.getCellStyle());
                cell.setCellValue(value.toString());
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }


}
