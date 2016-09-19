import java.util.List;

/**
 * Created by AndriiK on 9/19/2016.
 */
public class PoiJsonToXlsxRow {
    private List<PoiJsonToXlsxCell> cells;
    private int offset;

    public List<PoiJsonToXlsxCell> getCells() {
        return cells;
    }

    public void setCells(List<PoiJsonToXlsxCell> cells) {
        this.cells = cells;
    }

    public void setOffset(int offset) {
        this.offset = offset;
    }

    public int getOffset() {
        return offset;
    }
}
