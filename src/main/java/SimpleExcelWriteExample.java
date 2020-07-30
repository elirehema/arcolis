import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * -- This file created by eli on 14/07/2020 for poixss
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
import models.Book;
import models.Language;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SimpleExcelWriteExample {
    public static void main(String [] args) throws IOException {
        NiceExcelWriterExample excelWriterExample = new NiceExcelWriterExample();

        List<Language> languageList = excelWriterExample.getProgramingLanguage();
        String excelFilePath = "NiceJavaBooks.xlsx";
        String multipleFilePath = "BookList.xls";
        excelWriterExample.writeExcel(excelWriterExample.getListBook(), multipleFilePath);
        excelWriterExample.writeMultipleSheetExcel(languageList, excelFilePath);
        System.exit(0);
    }

    public static <T> List<T> myMethod(List<T> input) {
        return input; // just an example
    }

}
