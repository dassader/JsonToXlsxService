import java.util.ArrayList;
import java.util.List;

public class ExcelRow {
    private List<ExcelCell> excelCells;

    public ExcelRow() {
        this.excelCells = new ArrayList<ExcelCell>();
    }

    public List<ExcelCell> getExcelCells() {
        return excelCells;
    }

    public void setExcelCells(List<ExcelCell> excelCells) {
        this.excelCells = excelCells;
    }
}
