import java.util.List;

/**
 * Created by AndriiK on 9/19/2016.
 */
public class PoiJsonToXlsxCell {
    private String value;
    private int size;
    private int offset;
    private int colspan;

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    public int getOffset() {
        return offset;
    }

    public void setOffset(int offset) {
        this.offset = offset;
    }

    public int getColspan() {
        return colspan;
    }

    public void setColspan(int colspan) {
        this.colspan = colspan;
    }
}
