package ru.alfabank.urfu.brigadadena.excel.converter;

import org.apache.poi.ss.usermodel.CellCopyContext;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.alfabank.urfu.brigadadena.excel.columnsmodifiets.ColumnsModifier;
import ru.alfabank.urfu.brigadadena.excel.util.CellCopyUtils;

import java.io.Closeable;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

public class ExcelConverter implements Closeable {
    private final Workbook sourceWorkbook;
    private final Workbook sampleWorkbook;
    private final Workbook resultWorkbook;
    private final ProcessingSheets sheets;

    private final CellCopyContext copyContext = new CellCopyContext();
    private final List<ColumnsModifier> modifiers = new ArrayList<>();

    public ExcelConverter(Workbook source, Workbook result) {
        this.sourceWorkbook = source;
        checkCorrectWorkbook(this.sourceWorkbook);
        this.sampleWorkbook = result;
        checkCorrectWorkbook(this.sampleWorkbook);
        this.resultWorkbook = createResultWorkbook();

        this.sheets = new ProcessingSheets(
            this.sourceWorkbook.getSheetAt(0),
            this.sampleWorkbook.getSheetAt(0),
            this.resultWorkbook.getSheetAt(0)
        );
    }

    public void addModifier(ColumnsModifier modifier) {
        modifiers.add(modifier);
    }

    public Workbook getResult() {
        return getResult(sheets.source().getLastRowNum());
    }

    public Workbook getResult(int rowsCount) {
        createRows(rowsCount);
        for (var modifier : modifiers)
            modifier.execute(sheets, rowsCount, copyContext);
        autoSizeColumns();

        return resultWorkbook;
    }

    private void createRows(int rowsCount) {
        IntStream.rangeClosed(1, rowsCount)
            .parallel()
            .forEach(sheets.result()::createRow);
    }

    private void autoSizeColumns() {
        IntStream.range(0, sheets.result().getRow(0).getLastCellNum())
            .parallel()
            .forEach(sheets.result()::autoSizeColumn);
    }

    private Workbook createResultWorkbook() {
        var sampleSheet = this.sampleWorkbook.getSheetAt(0);

        Workbook result = new XSSFWorkbook();
        var sheet = result.createSheet(sampleSheet.getSheetName());
        var resultRow = sheet.createRow(0);
        for (var cell : sampleSheet.getRow(0)) {
            var resultCell = resultRow.createCell(cell.getColumnIndex(), cell.getCellType());

            CellCopyUtils.copyCellStyle(cell, resultCell, copyContext);
            CellCopyUtils.copyCellValue(cell, resultCell);
        }

        return result;
    }

    private void checkCorrectWorkbook(Workbook workbook) {
        if (workbook.getNumberOfSheets() == 0)
            throw new RuntimeException(); //TODO нормальные исключения
        var sheet = workbook.getSheetAt(0);
        if (sheet.getPhysicalNumberOfRows() < 2)
            throw new RuntimeException(); //TODO нормальные исключения
        if (sheet.getRow(0).getPhysicalNumberOfCells() < 1)
            throw new RuntimeException(); //TODO нормальные исключения
    }

    @Override
    public void close() throws IOException {
        sourceWorkbook.close();
        sampleWorkbook.close();
        resultWorkbook.close();
    }
}
