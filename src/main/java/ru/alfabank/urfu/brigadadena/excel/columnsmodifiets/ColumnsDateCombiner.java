package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.*;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

import java.util.Arrays;

public class ColumnsDateCombiner extends ColumnsModifier {
    private final int[] sourceColumnNums;
    private final int sampleColumnNum;

    public ColumnsDateCombiner(int[] sourceColumnNums, int sampleColumnNum) {
        this.sourceColumnNums = sourceColumnNums;
        this.sampleColumnNum = sampleColumnNum;
    }

    @Override
    protected void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext) {
        var sampleCell = sheets.sample().getRow(rowNum).getCell(sampleColumnNum);
        if (sampleCell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(sampleCell))
            throw new IllegalArgumentException(); // TODO адекватные исключения

        var resultCell = sheets.result().getRow(rowNum).createCell(sampleColumnNum, sampleCell.getCellType());
        CellCopyUtils.copyCellStyle(sampleCell, resultCell, copyContext);
        resultCell.setCellValue(getResultDate(sheets.source().getRow(rowNum)));
    }

    private double getResultDate(Row sourceRow) {
        return Arrays.stream(sourceColumnNums)
            .mapToObj(sourceRow::getCell)
            .mapToDouble(ColumnsDateCombiner::tryGetNumericDate)
            .sum();
    }

    private static double tryGetNumericDate(Cell cell) {
        //TODO Создать адекватные исключения
        if (cell.getCellType() != CellType.NUMERIC || !DateUtil.isCellDateFormatted(cell))
            throw new IllegalArgumentException();

        return cell.getNumericCellValue();
    }
}
