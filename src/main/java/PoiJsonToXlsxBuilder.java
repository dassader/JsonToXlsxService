import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.output.ByteArrayOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Created by AndriiK on 9/19/2016.
 */
public class PoiJsonToXlsxBuilder implements JsonToXlsxBuilder {

    private ObjectMapper objectMapper = new ObjectMapper();

    public InputStream build(String jsonData) {
        Workbook workbook = new SXSSFWorkbook();

        List<PoiJsonToXlsxSheet> poiJsonToXlsxSheets;
        try {
            poiJsonToXlsxSheets = objectMapper.readValue(jsonData, new TypeReference<List<PoiJsonToXlsxSheet>>(){});
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }

        for (PoiJsonToXlsxSheet poiJsonToXlsxSheet : poiJsonToXlsxSheets) {
            buildSheet(workbook, poiJsonToXlsxSheet);
        }

        //TODO: rename baos
        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try {
            workbook.write(baos);
        } catch (IOException e) {
            //TODO: rename
            throw new IllegalStateException(e);
        }

        return baos.toInputStream();
    }

    private void buildSheet(Workbook workbook, PoiJsonToXlsxSheet poiJsonToXlsxSheet) {
        String sheetName = poiJsonToXlsxSheet.getName();
        Sheet sheet;

        if(sheetName != null) {
            sheet = workbook.createSheet(sheetName);
        } else {
            sheet = workbook.createSheet();
        }

        List<PoiJsonToXlsxRow> rows = poiJsonToXlsxSheet.getRows();
        int position = 0;
        for (int i = 0; i < rows.size(); i++, position++) {
            position += rows.get(i).getOffset();
            buildRow(sheet, rows.get(i), position);
        }
    }

    private void buildRow(Sheet sheet, PoiJsonToXlsxRow poiJsonToXlsxRow, int rowPosition) {
        Row row = sheet.createRow(rowPosition);

        List<PoiJsonToXlsxCell> cells = poiJsonToXlsxRow.getCells();
        int position = 0;
        for (int i = 0; i < cells.size(); i++, position++) {
            position += cells.get(i).getOffset();
            buildCell(row, cells.get(i), position);
        }
    }

    private void buildCell(Row row, PoiJsonToXlsxCell poiJsonToXlsxCell, int cellPosition) {
        Cell cell = row.createCell(cellPosition);
        cell.setCellValue(poiJsonToXlsxCell.getValue());
    }
}
