import java.util.List;

public class JsonSheet {
    private String name;
    private List<JsonRow> rows;

    public List<JsonRow> getRows() {
        return rows;
    }

    public void setRows(List<JsonRow> rows) {
        this.rows = rows;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
