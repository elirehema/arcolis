import java.io.IOException;
import java.util.*;

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
import examples.ESheet;
import models.Book;
import styles.EType;

public class SimpleExcelWriteExample {
    public static void main(String [] args) throws IOException {
        NiceExcelWriterExample excelWriterExample = new NiceExcelWriterExample().getInstance();
        ESheet eSheet;

        List<Map<String, List<?>> > languageList = getListOfObject();
        String excelFilePath = "NiceJavaBooks";
        String multipleFilePath = "BookList";
        eSheet = new ESheet.Builder(multipleFilePath)

                .title(EType.TEAL_HEADER)
                .header(EType.FORMULA_1)
                .cell(EType.DEFAULT_CELL)
                .setMap(languageList)
                .writes();

        //excelWriterExample.ExcelSheet(getListBook("Book Name 1"), multipleFilePath);
        //excelWriterExample.ExcelSheets(languageList, excelFilePath);
        System.exit(0);
    }

    public static <T> List<T> myMethod(List<T> input) {
        return input; // just an example
    }

    public static List<Book> getListBook(String bookName) {
        Book book = null;

        Book book2 = new Book(bookName, "Joshua Bloch", 36, "Effective Java", "overflow.com",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");
        Book book3 = new Book(bookName, "Robert Martin", 42, "Effective Java", "overflow.com",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");
        Book book4 = new Book(bookName, "Bruce Eckel", 35, "Effective Java", "overflow",
                "Learn Java ", "Baron Quinn", "M8S4DS3211");

        List<Book> listBook = Arrays.asList(book2, book3, book4);
        List<Book> bookList = new ArrayList<>();
        for (int i = 0; i <= 100; i++) {
            book = new Book(bookName, "Kathy Bruce", 79 + i, "Effective Java", "overflow",
                    "Learn Java ", "Baron Quinn", "M8S4DS3211");
            bookList.add(book);
        }
        bookList.add(book4);
        bookList.add(book3);
        bookList.add(book2);


        return bookList;
    }
    public static List<Map<String, List<?>> > getListOfObject(){
        List<Map<String, List<?>>> map =  new ArrayList<Map<String, List<?>>>();
        Map<String, List<?>> map1 = new HashMap<>();
        map1.put("Kotlin", getListBook("Kotlin Essential"));
        map1.put("Java ", getListBook("Effective Java"));
        map1.put("Python", getListBook("Python for beginners"));
        map1.put("Javascript", getListBook("Javascript for Geeks"));

        map1.put("VueJs", getListBook("VueJs for Geeks"));

        map1.put("Ruby", getListBook("Ruby build the world"));
        map.add(map1);

        return map;
    }

    /*
    public static List<Language> getProgramingLanguage() {
        Language language1 = new Language("Java", getListBook());
        Language language2 = new Language("Python", getListBook());
        Language language3 = new Language("Javascript", getListBook());
        Language language4 = new Language("Ruby", getListBook());
        Language language5 = new Language("Kotlin", getListBook());


        List<Language> languageList = Arrays.asList(language1, language2, language3, language4, language5);

        return languageList;
    }
    */

}
