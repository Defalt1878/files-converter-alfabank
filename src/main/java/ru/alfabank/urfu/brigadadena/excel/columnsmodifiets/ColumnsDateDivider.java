package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

public class ColumnsDateDivider extends ColumnsModifier {
    private final int sourceColumnNum;
    private final int[] sampleColumnNums;

    public ColumnsDateDivider(int sourceColumnNum, int[] sampleColumnNums) {
        this.sourceColumnNum = sourceColumnNum;
        this.sampleColumnNums = sampleColumnNums;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sourceCell = sheets.source().getRow(rowNum).getCell(sourceColumnNum);
        if (sourceCell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(sourceCell))
            throw new IllegalArgumentException(); //TODO адекватные исключения

        var date = sourceCell.getNumericCellValue();

        for (var columnNum : sampleColumnNums) {
            var sampleCell = sheets.sample().getRow(rowNum).getCell(columnNum);
            if (sampleCell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(sampleCell))
                throw new IllegalArgumentException();

            var resultCell = sheets.result().getRow(rowNum).createCell(columnNum);
            CellCopyUtils.copyCellStyle(sampleCell, resultCell, copyContext);
            resultCell.setCellValue(date);
        }
    }
}
