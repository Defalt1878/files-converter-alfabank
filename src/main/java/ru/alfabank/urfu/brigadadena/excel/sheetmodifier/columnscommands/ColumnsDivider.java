package ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands;

import org.apache.poi.ss.usermodel.Sheet;
import ru.alfabank.urfu.brigadadena.excel.util.ExcelHelper;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

public class ColumnsDivider extends ColumnsModifier {
    private final int columnNum;
    private final String splitter;
    private int[] newColumnsNum;

    public ColumnsDivider(int columnNum, String splitter) {
        this.columnNum = columnNum;
        this.splitter = splitter;
    }

    @Override
    public int[] getColumnNums() {
        return new int[]{columnNum};
    }

    @Override
    public int[] getNewColumnNums() {
        return newColumnsNum == null ? new int[0] : newColumnsNum;
    }

    @Override
    public void apply(Sheet sheet) {
        var rowsData = getRowsData(sheet);
        if (newColumnsNum.length < 2)
            return;

        sheet.shiftColumns(columnNum + 1, sheet.getRow(0).getLastCellNum(), newColumnsNum.length - 1);
        setHeader(sheet);
        setData(sheet, rowsData);
    }

    private Map<Integer, String[]> getRowsData(Sheet sheet) {
        var lengths = new HashSet<Integer>();
        var result = new HashMap<Integer, String[]>();
        var skipFirst = true;
        for (var row : sheet) {
            if (skipFirst) {
                skipFirst = false;
                continue;
            }

            var data = ExcelHelper.getCellStringValue(row.getCell(columnNum)).split(Pattern.quote(splitter));
            result.put(row.getRowNum(), data);
            lengths.add(data.length);
        }

        if (lengths.size() == 0)
            throw new IllegalArgumentException("No data provided!"); //TODO адекватные исключения

        if (lengths.size() > 1)
            throw new IllegalArgumentException(); //TODO адекватные исключения

        newColumnsNum = IntStream.range(columnNum, columnNum + lengths.iterator().next()).toArray();

        return result;
    }

    private void setHeader(Sheet sheet) {
        var header = sheet.getRow(0);
        var columnName = ExcelHelper.getCellStringValue(header.getCell(columnNum));
        var columnNames = columnName.split(" \\+ ");
        if (columnNames.length != newColumnsNum.length)
            columnNames = IntStream.rangeClosed(1, newColumnsNum.length)
                .mapToObj(i -> String.format("%s [%s]", columnName, i))
                .toArray(String[]::new);

        for (var i = 0; i < newColumnsNum.length; i++) {
            var cell = header.createCell(columnNum + i);
            cell.setCellValue(columnNames[i]);
        }
    }

    private void setData(Sheet sheet, Map<Integer, String[]> rowsData) {
        for (var entry : rowsData.entrySet()) {
            var row = sheet.getRow(entry.getKey());
            var data = entry.getValue();

            for (var i = 0; i < newColumnsNum.length; i++) {
                var cell = row.createCell(newColumnsNum[i]);
                cell.setCellValue(data[i]);
            }
        }
    }
}
