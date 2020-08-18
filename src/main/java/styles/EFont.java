package styles;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

public class EFont {
    public EFont(){ }
    public Map<String, Font> createFont(Workbook workbook){
        Map<String, Font> fonts = new HashMap<>();
        Font font = workbook.createFont();

        font.setFontHeightInPoints((short)11);
        font.setColor(IndexedColors.WHITE.getIndex());
        fonts.put("default-white", font);

        font.setFontHeightInPoints((short)11);
        font.setColor(IndexedColors.BLACK.getIndex());
        fonts.put("default-black", font);

        font.setFontName("Operator Mono Medium");
        font.setColor(IndexedColors.LIGHT_GREEN.getIndex());
        fonts.put("operator-mono-light-green", font);

        font.setBold(true);
        font.setFontName("Operator Mono Medium");
        font.setFontHeightInPoints((short) 15);
        font.setColor(IndexedColors.WHITE.getIndex());
        fonts.put("operator-mono-bold-white", font);

        return fonts;
    }
}
