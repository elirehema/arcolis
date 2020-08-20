package examples;

import styles.EType;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public class ESheet {
    private List<?> objects;
    private List<Map<String, List<?>>> map;
    private String path;
    private EType title;
    private EType cell;
    private EType header;
    private EType background;


    public ESheet(List<?> objects, List<Map<String, List<?>>> map, String path, EType header, EType title, EType cell, EType background) {
        this.objects = objects;
        this.path = path;
        this.title = title;
        this.cell = cell;
        this.header = header;
        this.background = background;
        this.map = map;
    }
    public ESheet(List<Map<String, List<?>>> map, String path, EType header, EType title, EType cell, EType background) {
        this.path = path;
        this.title = title;
        this.cell = cell;
        this.header = header;
        this.background = background;
        this.map = map;
    }



    public EType getHeader() {
        return header;
    }

    public List<?> getObjects() {
        return objects;
    }

    public String getpath() {
        return path;
    }

    public EType gettitle() {
        return title;
    }

    public EType getcell() {
        return cell;
    }

    public EType getBackground() {
        return background;
    }

    public List<Map<String, List<?>>> getMap() {
        return map;
    }

    public static class Builder {
        private List<?> objects;
        List<Map<String, List<?>>> map;
        private final String path;
        private EType title;
        private EType cell;
        private EType header;
        private EType background;
        private EPrinter EPrinter = new EPrinter().getInstance();

        public Builder(List<?> objects, List<Map<String, List<?>>> map,
                       String path, EType header, EType title, EType cell, EType background) {
            this.objects = objects;
            this.path = path;
            this.title = title;
            this.cell = cell;
            this.header = header;
            this.background = background;
            this.map = map;
        }

        public Builder(String path) {
            this.path = path;
        }

        public Builder(List<?> objects, String path) {
            this.objects = objects;
            this.path = path;
        }

        public Builder(String path, List<Map<String, List<?>>> map) {
            this.map = map;
            this.path = path;
        }

        public Builder setMap(List<Map<String, List<?>>> map) {
            this.map = map;
            return this;
        }

        public Builder setData(List<?> objects) {
            this.objects = objects;
            return this;
        }

        public Builder title(EType title) {
            this.title = title;
            return this;
        }

        public Builder cell(EType cell) {
            this.cell = cell;
            return this;
        }

        public Builder header(EType header) {
            this.header = header;
            return this;
        }

        public Builder background(EType background) {
            this.background = background;
            return this;
        }

        public ESheet build() {
            return new ESheet(objects, null, path, header, title, cell, background);
        }

        public ESheet write() throws IOException {
            EPrinter.ExcelSheet(this.objects, path, header, title, cell);
            return new ESheet(objects, null, path, header, title, cell, background);
        }

        public ESheet writes() throws IOException {
            EPrinter.ExcelSheets(this.map , path, header, title, cell);
            return new ESheet( map, path, header, title, cell, background);
        }


    }
}