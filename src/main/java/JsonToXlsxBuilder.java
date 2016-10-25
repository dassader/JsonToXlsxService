import java.util.List;

public interface JsonToXlsxBuilder {
    byte[] build(List<JsonSheet> sheetList);
}
