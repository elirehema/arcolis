package examples;

import styles.EType;

import java.io.IOException;
import java.util.List;

public class ESheet {
    private final List<?> objects;
    private final String excelFilePath;
    private final EType title;
    private final EType cell;
    private final EType header;


    public ESheet(List<?> objects, String excelFilePath,EType header, EType title, EType cell) {
        this.objects = objects;
        this.excelFilePath = excelFilePath;
        this.title = title;
        this.cell = cell;
        this.header = header;
    }

    public EType getHeader() {
        return header;
    }

    public List<?> getObjects() {
        return objects;
    }

    public String getExcelFilePath() {
        return excelFilePath;
    }

    public EType gettitle() {
        return title;
    }

    public EType getcell() {
        return cell;
    }

    public static class Builder {
        private List<?> objects;
        private final String excelFilePath;
        private  EType title;
        private EType cell;
        private EType header;
        private EPrinter EPrinter = new EPrinter().getInstance();

        public Builder(List<?> objects, String excelFilePath, EType header, EType title, EType cell) {
            this.objects = objects;
            this.excelFilePath = excelFilePath;
            this.title = title;
            this.cell = cell;
            this.header = header;
        }
        public Builder(String excelFilePath){
            this.excelFilePath = excelFilePath;
        }
        public Builder(List<?> objects, String excelFilePath){
            this.objects = objects;
            this.excelFilePath = excelFilePath;
        }
        public Builder setData(List<?> objects){
            this.objects = objects;
            return this;
        }

        public Builder title(EType title){
            this.title = title;
            return this;
        }
        public Builder cell(EType cell){
            this.cell = cell;
            return this;
        }
        public Builder header(EType header){
            this.header = header;
            return this;
        }

        public ESheet build() {
            return new ESheet(objects, excelFilePath,header, title,cell);
        }
        public ESheet write() throws IOException {
            EPrinter.ExcelSheet(this.objects, excelFilePath, header, title, cell);

            return new ESheet(objects,excelFilePath,header, title,cell);
        }

    }
}