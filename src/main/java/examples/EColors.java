package examples;

import org.apache.poi.ss.usermodel.IndexedColors;
import styles.EType;

public class EColors {

    private IndexedColors indexedColors;
    private EType eType;

    public short getIndexedColors() {
        return indexedColors.index;
    }

    public void setIndexedColors(IndexedColors indexedColors) {
        this.indexedColors = indexedColors;
    }

    public EType geteType() {
        return eType;
    }

    public void seteType(EType eType) {
        this.eType = eType;
    }

    public EColors(IndexedColors indexedColors, EType eType) {
        this.indexedColors = indexedColors;
        this.eType = eType;
    }


}
