import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.commons.io.output.ByteArrayOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
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
            poiJsonToXlsxSheets = objectMapper.readValue(jsonData, new TypeReference<List<PoiJsonToXlsxSheet>>() {
            });
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

        if (sheetName != null) {
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

        setColumnAutoSize(sheet);
    }

    private void setColumnAutoSize(Sheet sheet) {
        Iterator<Row> rowIterator = sheet.rowIterator();
        int maxCellNum = 0;
        while (rowIterator.hasNext()) {
            Row targetRow = rowIterator.next();
            int lastCellNum = targetRow.getLastCellNum();

            if(lastCellNum > maxCellNum) {
                maxCellNum = lastCellNum;
            }
        }

        for (int i = 0; i < maxCellNum; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private void buildRow(Sheet sheet, PoiJsonToXlsxRow poiJsonToXlsxRow, int rowPosition) {
        Row row = sheet.createRow(rowPosition);

        List<PoiJsonToXlsxCell> cells = poiJsonToXlsxRow.getCells();
        int cellPosition = 0;
        for (int i = 0; i < cells.size(); i++, cellPosition++) {
            PoiJsonToXlsxCell targetCell = cells.get(i);
            int colspanValue = targetCell.getColspan();

            cellPosition += targetCell.getOffset();

            buildCell(row, targetCell, cellPosition);
            handleColspan(sheet, colspanValue, rowPosition, cellPosition);

            if(colspanValue != 0) {
                cellPosition += colspanValue - 1;
            }
        }
    }

    private void handleColspan(Sheet sheet, int colspan, int rowPosition, int cellPosition) {
        if(colspan != 0) {
            int lastCol = cellPosition + colspan - 1;
            sheet.addMergedRegion(new CellRangeAddress(rowPosition, rowPosition, cellPosition, lastCol));
        }
    }

    private void buildCell(Row row, PoiJsonToXlsxCell poiJsonToXlsxCell, int cellPosition) {
        Cell cell = row.createCell(cellPosition);
        cell.setCellValue(poiJsonToXlsxCell.getValue());
    }
}
