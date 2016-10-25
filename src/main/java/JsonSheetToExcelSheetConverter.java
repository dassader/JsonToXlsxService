import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class JsonSheetToExcelSheetConverter {
    public ExcelSheet convert(JsonSheet jsonSheet) {
        ExcelSheet excelSheet = new ExcelSheet(jsonSheet.getName());

        List<JsonRow> jsonRows = jsonSheet.getRows();

        int rowNumber = 0;

        Map<Integer, List<Interval>> offsets = new HashMap<Integer, List<Interval>>();

        List<ExcelRow> excelRows = excelSheet.getExcelRows();

        for (int i = 0; i < jsonRows.size(); i++, rowNumber++) {
            JsonRow jsonRow = jsonRows.get(rowNumber);

            excelRows.add(handleRow(rowNumber, jsonRow, offsets));
        }

        return excelSheet;
    }

    private ExcelRow handleRow(int rowNumber, JsonRow jsonRow, Map<Integer, List<Interval>> offsets) {
        List<JsonCell> jsonCells = jsonRow.getCells();

        ExcelRow excelRow = new ExcelRow();

        List<ExcelCell> excelCells = excelRow.getExcelCells();

        int cellNumber = 0;

        for (int i = 0; i < jsonCells.size(); i++, cellNumber++) {
            JsonCell jsonCell = jsonCells.get(cellNumber);

            excelCells.add(handleCell(rowNumber, cellNumber, jsonCell, offsets));
        }

        return excelRow;
    }

    private ExcelCell handleCell(int rowNumber, int cellNumber, JsonCell jsonCell, Map<Integer, List<Interval>> offsets) {
        int fromCell = calculateFromCell(rowNumber, cellNumber, offsets);
        int toCell = fromCell + jsonCell.getColspan();

        if(jsonCell.getColspan() != 0) {
            toCell -= 1;
        }

        int toRow = rowNumber + jsonCell.getRowspan();

        if(jsonCell.getRowspan() != 0) {
            toRow -= 1;
        }

        ExcelCell excelCell = new ExcelCell(fromCell, toCell, rowNumber, toRow, jsonCell.getValue());

        updateOffsets(excelCell, offsets);

        return excelCell;
    }

    private void updateOffsets(ExcelCell excelCell, Map<Integer, List<Interval>> offsets) {
        for (int i = excelCell.getFromRow(); i <= excelCell.getToRow(); i++) {
            List<Interval> intervals = offsets.get(i);

            if(intervals == null) {
                intervals = new ArrayList<Interval>();
                offsets.put(i, intervals);
            }

            intervals.add(new Interval(excelCell.getFromCell(), excelCell.getToCell()));
        }
    }

    private int calculateFromCell(int rowNumber, int cellNumber, Map<Integer, List<Interval>> offsets) {
        List<Interval> intervals = offsets.get(rowNumber);

        if(intervals == null) {
            return cellNumber;
        }

        for (Interval interval : intervals) {
            if (isInclude(cellNumber, interval)) {
                return calculateFromCell(rowNumber, interval.getTo() + 1, offsets);
            }
        }

        return cellNumber;
    }

    private boolean isInclude(int rowNumber, Interval entry) {
        return rowNumber >= entry.getFrom() && rowNumber <= entry.getTo();
    }

    private static class Interval {
        private int from;
        private int to;

        public Interval(int from, int to) {
            this.from = from;
            this.to = to;
        }

        public int getFrom() {
            return from;
        }

        public void setFrom(int from) {
            this.from = from;
        }

        public int getTo() {
            return to;
        }

        public void setTo(int to) {
            this.to = to;
        }
    }
}
