package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

import java.util.Arrays;
import java.util.Locale;
import java.util.stream.Collectors;

public class ColumnsSimpleCombiner extends ColumnsModifier {
    private final int[] sourceColumnNums;
    private final int sampleColumnNum;
    private final String splitter;

    private final DataFormatter dataFormatter = new DataFormatter(Locale.getDefault());

    public ColumnsSimpleCombiner(int[] sourceColumnNums, int sampleColumnNum, String splitter) {
        this.sourceColumnNums = sourceColumnNums;
        this.sampleColumnNum = sampleColumnNum;
        this.splitter = splitter;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sampleCell = sheets.sample().getRow(rowNum).getCell(sampleColumnNum);
        var resultCell = sheets.result().getRow(rowNum).createCell(sampleColumnNum, sampleCell.getCellType());

        CellCopyUtils.copyCellStyle(sampleCell, resultCell, copyContext);
        resultCell.setCellValue(getSimpleResultData(sheets.source().getRow(rowNum)));
    }

    private String getSimpleResultData(Row sourceRow) {
        return Arrays.stream(sourceColumnNums)
            .mapToObj(sourceRow::getCell)
            .map(dataFormatter::formatCellValue)
            .collect(Collectors.joining(splitter));
    }
}
