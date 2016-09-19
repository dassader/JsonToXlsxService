import java.util.List;

public class PoiJsonToXlsxSheet {
    private String name;
    private List<PoiJsonToXlsxRow> rows;

    public List<PoiJsonToXlsxRow> getRows() {
        return rows;
    }

    public void setRows(List<PoiJsonToXlsxRow> rows) {
        this.rows = rows;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
