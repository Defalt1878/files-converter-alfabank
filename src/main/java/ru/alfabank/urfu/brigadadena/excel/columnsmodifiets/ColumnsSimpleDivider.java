package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.DataFormatter;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

import java.util.Locale;

public class ColumnsSimpleDivider extends ColumnsModifier {
    private final int sourceColumnNum;
    private final int[] sampleColumnNums;
    private final String splitter;

    private final DataFormatter dataFormatter = new DataFormatter(Locale.getDefault());

    public ColumnsSimpleDivider(int sourceColumnNum, int[] sampleColumnNums, String splitter) {
        this.sourceColumnNum = sourceColumnNum;
        this.sampleColumnNums = sampleColumnNums;
        this.splitter = splitter;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sourceCell = sheets.source().getRow(rowNum).getCell(sourceColumnNum);
        var data = dataFormatter.formatCellValue(sourceCell).split(splitter);
        if (sampleColumnNums.length != data.length)
            throw new RuntimeException(); //TODO адекватные исключения

        for (var i = 0; i < data.length; i++) {
            var sampleCell = sheets.sample().getRow(rowNum).getCell(sampleColumnNums[i]);
            var resultCell = sheets.result().getRow(rowNum).createCell(sampleColumnNums[i], sampleCell.getCellType());
            CellCopyUtils.copyCellStyle(sampleCell, resultCell, copyContext);
            resultCell.setCellValue(data[i]);
        }
    }
}
