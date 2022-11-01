package ru.alfabank.urfu.brigadadena.excel.sheetConnector;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import ru.alfabank.urfu.brigadadena.excel.util.ExcelHelper;

import java.util.stream.IntStream;

public class ColumnsConnector {
    private final int srcColumnNum;
    private final int dstColumnNum;
    private final CellStyle dstStyle;

    public int getSrcColumnNum() {
        return srcColumnNum;
    }

    public int getDstColumnNum() {
        return dstColumnNum;
    }

    public CellStyle getDstStyle() {
        return dstStyle;
    }

    public ColumnsConnector(int srcColumnNum, int dstColumnNum, CellStyle dstStyle) {
        this.srcColumnNum = srcColumnNum;
        this.dstColumnNum = dstColumnNum;
        this.dstStyle = dstStyle;
    }

    public void apply(Sheet src, Sheet dst, int rowsCount) {
        IntStream.rangeClosed(1, Math.min(rowsCount, src.getLastRowNum()))
            .parallel()
            .forEach(rowNum -> modify(src, dst, rowNum));
    }

    private void modify(Sheet src, Sheet dst, int rowNum) {
        var srcCell = src.getRow(rowNum).getCell(srcColumnNum);

        var dstRow = dst.getRow(rowNum);
        if (dstRow == null)
            dstRow = dst.createRow(rowNum);
        var dstCell = dstRow.createCell(dstColumnNum);

        dstCell.setCellStyle(dstStyle);
        ExcelHelper.copyCellValue(srcCell, dstCell);
    }
}
