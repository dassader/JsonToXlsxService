import java.util.List;

public class JsonRow {
    private List<JsonCell> cells;
    private int offset;

    public List<JsonCell> getCells() {
        return cells;
    }

    public void setCells(List<JsonCell> cells) {
        this.cells = cells;
    }

    public void setOffset(int offset) {
        this.offset = offset;
    }

    public int getOffset() {
        return offset;
    }
}
