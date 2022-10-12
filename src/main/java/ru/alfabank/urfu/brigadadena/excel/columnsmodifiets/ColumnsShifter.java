package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

public class ColumnsShifter extends ColumnsModifier {
    private final int sourceColumnNum;
    private final int sampleColumnNum;

    public ColumnsShifter(int sourceColumnNum, int sampleColumnNum) {
        this.sourceColumnNum = sourceColumnNum;
        this.sampleColumnNum = sampleColumnNum;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sourceCell = sheets.source().getRow(rowNum).getCell(sourceColumnNum);
        var sampleCell = sheets.sample().getRow(rowNum).getCell(sampleColumnNum);
        var resultCell = sheets.result().getRow(rowNum).createCell(sampleColumnNum, sampleCell.getCellType());

        CellCopyUtils.copyCellStyle(sourceCell, resultCell, copyContext);
        CellCopyUtils.copyCellValue(sourceCell, resultCell);
    }
}
