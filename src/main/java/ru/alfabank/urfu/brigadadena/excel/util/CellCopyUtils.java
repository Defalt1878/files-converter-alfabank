package ru.alfabank.urfu.brigadadena.excel.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.util.CellUtil;

public class CellCopyUtils {
    private static final CellCopyPolicy styleCopyPolicy;
    private static final CellCopyPolicy valueCopyPolicy;

    static {
        var emptyCopyPolicy = new CellCopyPolicy.Builder()
            .cellValue(false)
            .cellStyle(false)
            .cellFormula(false)
            .copyHyperlink(false)
            .mergeHyperlink(false)
            .rowHeight(false)
            .condenseRows(false)
            .mergedRegions(false)
            .build();

        styleCopyPolicy = emptyCopyPolicy.createBuilder()
            .cellStyle(true)
            .build();
        valueCopyPolicy = emptyCopyPolicy.createBuilder()
            .cellValue(true)
            .cellFormula(true)
            .build();
    }

    public static void copyCellValue(Cell src, Cell dst) {
        CellUtil.copyCell(src, dst, valueCopyPolicy, null);
    }

    public static void copyCellStyle(Cell src, Cell dst, CellCopyContext context) {
        CellUtil.copyCell(src, dst, styleCopyPolicy, context);
    }
}
