package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

public class ColumnsDateDivider extends ColumnsModifier {
    private final int sourceColumnNum;
    private final int sampleDateColumnNum;
    private final int sampleTimeColumnNum;

    public ColumnsDateDivider(int sourceColumnNum, int sampleDateColumnNum, int sampleTimeColumnNum) {
        this.sourceColumnNum = sourceColumnNum;
        this.sampleDateColumnNum = sampleDateColumnNum;
        this.sampleTimeColumnNum = sampleTimeColumnNum;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sourceCell = sheets.source().getRow(rowNum).getCell(sourceColumnNum);
        if (sourceCell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(sourceCell))
            throw new IllegalArgumentException(); //TODO адекватные исключения

        var dateTime = sourceCell.getLocalDateTimeCellValue();

        var resultDateCell = sheets.result().getRow(rowNum).createCell(sampleDateColumnNum);
        var sampleDateCell = sheets.sample().getRow(rowNum).getCell(sampleDateColumnNum);
        CellCopyUtils.copyCellStyle(sampleDateCell, resultDateCell, copyContext);
        resultDateCell.setCellValue(dateTime);

        var resultTimeCell = sheets.result().getRow(rowNum).createCell(sampleTimeColumnNum);
        var sampleTimeCell = sheets.sample().getRow(rowNum).getCell(sampleTimeColumnNum);
        CellCopyUtils.copyCellStyle(sampleTimeCell, resultTimeCell, copyContext);
        resultTimeCell.setCellValue(dateTime);
        //TODO отрефакторить
    }
}
