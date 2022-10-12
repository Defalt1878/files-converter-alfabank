package ru.alfabank.urfu.brigadadena.excel.columnsmodifiets;

import org.apache.poi.ss.usermodel.CellCopyContext;
import ru.alfabank.urfu.brigadadena.excel.converter.ProcessingSheets;

import java.util.stream.IntStream;

public abstract class ColumnsModifier {
    public void execute(ProcessingSheets sheets, int rowsCount, CellCopyContext copyContext) {
        IntStream.rangeClosed(1, rowsCount)
            .parallel()
            .forEach(rowNum -> modify(sheets, rowNum, copyContext));
    }

    protected abstract void modify(ProcessingSheets sheets, int rowNum, CellCopyContext copyContext);
}
