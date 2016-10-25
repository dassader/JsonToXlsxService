import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import static org.apache.poi.ss.usermodel.CellStyle.VERTICAL_TOP;

@Service
public class PoiJsonToExcelBuilder implements JsonToXlsxBuilder {
    @Autowired
    private JsonSheetToExcelSheetConverter sheetConverter;

    public byte[] build(List<JsonSheet> sheetList) {
        Workbook workbook = new SXSSFWorkbook();

        for (JsonSheet jsonSheet : sheetList) {
            buildSheet(workbook, sheetConverter.convert(jsonSheet));
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new IllegalStateException(e);
        }

        return outputStream.toByteArray();
    }

    private void buildSheet(Workbook workbook, ExcelSheet excelSheet) {
        Sheet sheet = buildSheet(excelSheet, workbook);

        List<ExcelRow> excelRows = excelSheet.getExcelRows();

        for (int rowPosition = 0; rowPosition < excelRows.size(); rowPosition++) {
            buildRow(rowPosition, excelRows.get(rowPosition), sheet, workbook);
        }

        setColumnAutoSize(sheet);
    }

    private Sheet buildSheet(ExcelSheet excelSheet, Workbook workbook) {
        String sheetName = excelSheet.getName();
        if (sheetName != null) {
            return workbook.createSheet(sheetName);
        } else {
            return workbook.createSheet();
        }
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

        applyAutoSize(sheet, maxCellNum);
    }

    private void applyAutoSize(Sheet sheet, int maxCellNum) {
        for (int i = 0; i < maxCellNum; i++) {
            sheet.autoSizeColumn(i, true);
        }
    }

    private void buildRow(int rowPosition, ExcelRow jsonRow, Sheet sheet, Workbook workbook) {
        Row row = sheet.createRow(rowPosition);

        for (ExcelCell excelCell : jsonRow.getExcelCells()) {
            buildCell(row, excelCell, sheet, workbook);
        }
    }

    private void buildCell(Row row, ExcelCell excelCell, Sheet sheet, Workbook workbook) {
        Cell cell = row.createCell(excelCell.getFromCell());

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setVerticalAlignment(VERTICAL_TOP);
        cell.setCellStyle(cellStyle);

        cell.setCellValue(excelCell.getText());
        sheet.addMergedRegion(new CellRangeAddress(excelCell.getFromRow(), excelCell.getToRow(),
                excelCell.getFromCell(), excelCell.getToCell()));
    }
}
