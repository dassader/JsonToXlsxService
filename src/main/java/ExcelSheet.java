import java.util.ArrayList;
import java.util.List;

public class ExcelSheet {
    private List<ExcelRow> excelRows;
    private String name;

    public ExcelSheet(String name) {
        this.name = name;
        this.excelRows = new ArrayList<ExcelRow>();
    }

    public List<ExcelRow> getExcelRows() {
        return excelRows;
    }

    public void setExcelRows(List<ExcelRow> excelRows) {
        this.excelRows = excelRows;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
