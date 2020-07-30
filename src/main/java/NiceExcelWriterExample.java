import cellstyles.ECellStyle;
import models.Book;
import models.Language;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import workbook.EWorkook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.Arrays;
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
    private ECellStyle eCellStyle = new ECellStyle();
    private EWorkook eWorkook = new EWorkook();

    /**
     * Write an Excel File with single Sheet
     **/
    public void writeExcel(List<Book> objectList, String excelFilePath) throws IOException {
        Workbook workbook = new HSSFWorkbook();
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

    public void writeMultipleSheetExcel(List<Language> languages, String excelFilePath) throws IOException {
        Workbook workbook = eWorkook.getWorkbook(excelFilePath);
        for (Language parentObject : languages) {
            Sheet sheet = workbook.createSheet(parentObject.getName());
            int rowCount = 0;
            Field[] fields = parentObject.getClass().getDeclaredFields();
            for (Field f : fields) {

                try {
                    Field field = parentObject.getClass().getDeclaredField(f.getName().toString());
                    field.setAccessible(true);
                    Object object = field.get(parentObject);

                    if (object instanceof Collection) {
                        System.out.println("Collection Instance");
                    }
                    if (object instanceof String) {
                        System.out.println("String Instance:  " + object.toString());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            for (Object obj : parentObject.getBooks()) {
                createHeaderRow(obj, sheet);
                Row row = sheet.createRow(++rowCount);
                try {
                    writeBook(obj, rowCount, row);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            rowCount = 0;
        }
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }
    }

    private void createHeaderRow(Object o, Sheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        PrintSetup printSetup = sheet.getPrintSetup();

        /**Set Page Number on Footer **/
        Footer ft = sheet.getFooter();
        ft.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        /**Set Fonts **/
        Font font = sheet.getWorkbook().createFont();
        font.setBoldweight((short) 12);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());
        cellStyle.setFont(font);

        /**Set Print Format**/
        sheet.setAutobreaks(true);
        printSetup.setFitHeight((short) 1);
        printSetup.setFitWidth((short) 1);

        /** Set Bottom Styles **/
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_DOTTED);
        cellStyle.setBottomBorderColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

        Row row = sheet.createRow(0);
        sheet.setDefaultColumnWidth(15);
        sheet.setFitToPage(true);

        Cell cellNumber = row.createCell(0);
        cellNumber.setCellStyle(cellStyle);
        cellNumber.setCellValue("#");
        int index = 0;
        for (Field f : o.getClass().getDeclaredFields()) {
            Cell cell = row.createCell(++index);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(f.getName().toUpperCase());
        }
    }

    private void writeBook(Object obj, Integer rowNumber, Row row) {
        Cell cell = row.createCell(0);
        CellStyle cellStyle = eCellStyle.getCellStyle(cell);

        CellStyle backgroundStyle = eCellStyle.getCellStyle(cell);
        backgroundStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        backgroundStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell.setCellValue(rowNumber);
        cell.setCellStyle(backgroundStyle);
        Field[] fields = obj.getClass().getDeclaredFields();
        int rowCount = 0;
        for (Field f : fields) {
            f.setAccessible(true);
            cell = row.createCell(++rowCount);
            try {
                Field field = obj.getClass().getDeclaredField(f.getName().toString());
                field.setAccessible(true);
                Object value = field.get(obj);
                cell.setCellValue(value.toString());
            } catch (Exception e) {
                e.printStackTrace();
            }
            cell.setCellStyle(cellStyle);
        }

    }




}
