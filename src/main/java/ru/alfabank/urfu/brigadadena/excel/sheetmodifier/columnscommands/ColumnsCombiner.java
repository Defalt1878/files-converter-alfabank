package ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands;

import org.apache.poi.ss.usermodel.Sheet;
import ru.alfabank.urfu.brigadadena.excel.util.ExcelHelper;

import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

public class ColumnsCombiner extends ColumnsModifier {
    private final int[] columnNums;
    private Integer newColumnNum = null;
    private final String splitter;

    public ColumnsCombiner(int[] columnNums, String splitter) {
        this.columnNums = columnNums;
        this.splitter = splitter;
    }

    @Override
    public int[] getColumnNums() {
        return columnNums;
    }

    @Override
    public int[] getNewColumnNums() {
        return newColumnNum == null ? new int[0] : new int[]{newColumnNum};
    }

    @Override
    public void apply(Sheet sheet) {
        if (columnNums.length < 2)
            return;
        var rowsData = getRowsData(sheet);

        setHeader(sheet);
        setData(sheet, rowsData);
        shiftColumns(sheet);
    }

    private Map<Integer, String> getRowsData(Sheet sheet) {
        var result = new HashMap<Integer, String>();
        StreamSupport.stream(sheet.spliterator(), false)
            .skip(1)
            .forEach(row -> result.put(
                         row.getRowNum(),
                         Arrays.stream(columnNums)
                             .mapToObj(row::getCell)
                             .map(ExcelHelper::getCellStringValue)
                             .collect(Collectors.joining(splitter))
                     )
            );

        newColumnNum = columnNums[0];

        return result;
    }

    private void setHeader(Sheet sheet) {
        var header = sheet.getRow(0);
        var columnNames = Arrays.stream(columnNums)
            .mapToObj(header::getCell)
            .map(ExcelHelper::getCellStringValue)
            .toList();
        var cell = header.createCell(newColumnNum);
        String resultName = null;
        if (columnNames.stream().allMatch(name -> name.matches(".* \\[\\d+]$"))) {
            var namesSet = columnNames.stream()
                .map(name -> name.split(" \\[\\d+]")[0])
                .collect(Collectors.toSet());
            if (namesSet.size() == 1)
                resultName = namesSet.iterator().next();
        }
        if (resultName == null)
            resultName = String.join(" + ", columnNames);

        cell.setCellValue(resultName);
    }

    private void setData(Sheet sheet, Map<Integer, String> rowsData) {
        for (var entry : rowsData.entrySet()) {
            var cell = sheet.getRow(entry.getKey()).createCell(newColumnNum);
            cell.setCellValue(entry.getValue());
        }
    }

    private void shiftColumns(Sheet sheet) {
        deleteOldCells(sheet);
        var sorted = Arrays.stream(columnNums)
            .skip(1)
            .map(i -> -i).sorted().map(i -> -i)
            .toArray();
        var shiftSize = 1;
        var lastModifiedColNum = sorted[0];
        for (var i = 1; i < sorted.length; i++) {
            var columnNum = sorted[i];
            if (lastModifiedColNum - shiftSize == columnNum) {
                shiftSize++;
                continue;
            }
            if (lastModifiedColNum < sheet.getRow(0).getLastCellNum())
                sheet.shiftColumns(lastModifiedColNum + 1, sheet.getRow(0).getLastCellNum(), -shiftSize);
            shiftSize = 1;
            lastModifiedColNum = columnNum;
        }

        if (lastModifiedColNum < sheet.getRow(0).getLastCellNum())
            sheet.shiftColumns(lastModifiedColNum + 1, sheet.getRow(0).getLastCellNum(), -shiftSize);
    }

    private void deleteOldCells(Sheet sheet) {
        for (var row : sheet) {
            for (var columnNum : columnNums) {
                if (columnNum == newColumnNum)
                    continue;
                var cell = row.getCell(columnNum);
                if (cell != null)
                    row.removeCell(cell);
            }
        }
    }
}
