package examples;

import styles.EType;

import java.io.IOException;
import java.util.Date;
import java.util.List;

public class ESheet {
    private final List<?> objects;
    private final String excelFilePath;
    private final EType headerCellColor;
    private final EType dataCellColor;


    public ESheet(List<?> objects, String excelFilePath, EType headerCellColor, EType dataCellColor) {
        this.objects = objects;
        this.excelFilePath = excelFilePath;
        this.headerCellColor = headerCellColor;
        this.dataCellColor = dataCellColor;
    }

    public List<?> getObjects() {
        return objects;
    }

    public String getExcelFilePath() {
        return excelFilePath;
    }

    public EType getheaderCellColor() {
        return headerCellColor;
    }

    public EType getdataCellColor() {
        return dataCellColor;
    }

    public static class Builder {
        private final List<?> objects;
        private final String excelFilePath;
        private  EType headerCellColor;
        private EType dataCellColor;
        private  Aerosol aerosol= new Aerosol().getInstance();

        public Builder(List<?> objects, String excelFilePath, EType headerCellColor, EType dataCellColor) {
            this.objects = objects;
            this.excelFilePath = excelFilePath;
            this.headerCellColor = headerCellColor;
            this.dataCellColor = dataCellColor;
        }
        public Builder(List<?> objects, String excelFilePath){
            this.objects = objects;
            this.excelFilePath = excelFilePath;

        }

        public Builder headerCellColor(EType headerCellColor){
            this.headerCellColor = headerCellColor;
            return this;
        }
        public Builder dataCellColor(EType dataCellColor){
            this.dataCellColor = dataCellColor;
            return this;
        }

        public ESheet build() {
            return new ESheet(objects, excelFilePath,headerCellColor,dataCellColor);
        }
        public ESheet write() throws IOException {
            aerosol.ExcelSheet(
                    this.objects, excelFilePath,headerCellColor, dataCellColor
            );

            return new ESheet(objects,excelFilePath,headerCellColor,dataCellColor);
        }

    }
}