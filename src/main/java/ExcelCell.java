public class ExcelCell {
    private int fromCell;
    private int toCell;
    private int fromRow;
    private int toRow;
    private String text;

    public ExcelCell(int fromCell, int toCell, int fromRow, int toRow, String text) {
        this.fromCell = fromCell;
        this.toCell = toCell;
        this.fromRow = fromRow;
        this.toRow = toRow;
        this.text = text;
    }

    public int getFromCell() {
        return fromCell;
    }

    public void setFromCell(int fromCell) {
        this.fromCell = fromCell;
    }

    public int getToCell() {
        return toCell;
    }

    public void setToCell(int toCell) {
        this.toCell = toCell;
    }

    public int getFromRow() {
        return fromRow;
    }

    public void setFromRow(int fromRow) {
        this.fromRow = fromRow;
    }

    public int getToRow() {
        return toRow;
    }

    public void setToRow(int toRow) {
        this.toRow = toRow;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }
}
