package ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands;

import org.apache.poi.ss.usermodel.Sheet;

public abstract class ColumnsModifier {
    public abstract int[] getColumnNums();
    public abstract int[] getNewColumnNums();

    public abstract void apply(Sheet sheet);
}
