package ru.alfabank.urfu.brigadadena.excel.converter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.alfabank.urfu.brigadadena.excel.sheetConnector.ColumnsConnector;
import ru.alfabank.urfu.brigadadena.excel.sheetConnector.SheetConnector;
import ru.alfabank.urfu.brigadadena.excel.sheetmodifier.SheetModifier;
import ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands.ColumnsCombiner;
import ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands.ColumnsDivider;
import ru.alfabank.urfu.brigadadena.excel.sheetmodifier.columnscommands.ColumnsModifier;

import java.io.Closeable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.stream.StreamSupport;

public class ExcelConverter implements Closeable {
    private final Workbook source;
    private final Workbook cutSource;
    private Workbook modifiableSource;
    private final Workbook result;
    private final CellStyle[] sampleStyles;
    private final SheetModifier sheetModifier = new SheetModifier();
    private final SheetConnector sheetConnector = new SheetConnector();

    public ExcelConverter(XSSFWorkbook source, XSSFWorkbook sample) throws IOException {
        checkCorrectWorkbook(source);
        this.source = source;
        checkCorrectWorkbook(sample);
        this.result = sample;
        this.cutSource = copyWorkbook(source);
        cutRows(this.cutSource.getSheetAt(0), 4);
        this.modifiableSource = copyWorkbook(cutSource);
        this.sampleStyles = getStyles(this.result.getSheetAt(0).getRow(1));
        cutRows(this.result.getSheetAt(0), 1);
    }

    public Workbook getSourceExample() {
        return modifiableSource;
    }

    public Workbook getResultExample() {
        return result;
    }

    public Workbook getFinalResult() {
        var sourceSheet = source.getSheetAt(0);
        sheetModifier.applyAll(sourceSheet);
        sheetConnector.applyAll(sourceSheet, result.getSheetAt(0), sourceSheet.getLastRowNum());
        return result;
    }

    public void combineColumns(int[] columnNums, String splitter) {
        handleModifier(new ColumnsCombiner(columnNums, splitter));
    }

    public void divideColumns(int columnNum, String splitter) {
        handleModifier(new ColumnsDivider(columnNum, splitter));
    }

    private void handleModifier(ColumnsModifier modifier) {
        sheetModifier.newModifier(modifier);
        sheetModifier.applyLastAdded(modifiableSource.getSheetAt(0));
        updateConnectors(modifier.getColumnNums());
    }

    public void cancelLast() throws IOException {
        modifiableSource = copyWorkbook(cutSource);
        var removed = sheetModifier.removeLast();
        updateConnectors(removed.getNewColumnNums());
    }

    private void updateConnectors(int[] changedColumnsNums) {
        if (sheetConnector.removeSrcIntersections(changedColumnsNums)) {
            cutRows(result.getSheetAt(0), 1);
            sheetConnector.applyAll(modifiableSource.getSheetAt(0), result.getSheetAt(0), 3);
        }
    }

    public void connectColumns(int srcColumnNum, int dstColumnNum) {
        sheetConnector.newConnector(new ColumnsConnector(srcColumnNum, dstColumnNum, sampleStyles[dstColumnNum]));
        sheetConnector.applyLast(modifiableSource.getSheetAt(0), result.getSheetAt(0), 3);
    }


    private Workbook copyWorkbook(Workbook source) throws IOException {
        final String temp = "temp.xlsx";
        XSSFWorkbook result;
        try (var outputStream = new FileOutputStream(temp)) {
            source.write(outputStream);
        }
        try (var inputStream = new FileInputStream(temp)) {
            result = new XSSFWorkbook(inputStream);
        }
        Files.delete(Paths.get(temp));
        return result;
    }

    private void cutRows(Sheet sheet, int cutStart) {
        for (var rowNum = cutStart; rowNum <= sheet.getLastRowNum(); rowNum++) {
            var row = sheet.getRow(rowNum);
            if (row != null)
                sheet.removeRow(row);
        }
    }

    private CellStyle[] getStyles(Row row) {
        return StreamSupport.stream(row.spliterator(), false)
            .map(Cell::getCellStyle)
            .toArray(CellStyle[]::new);
    }

    private void checkCorrectWorkbook(Workbook workbook) {
        if (workbook.getNumberOfSheets() == 0)
            throw new RuntimeException(); //TODO нормальные исключения
        var sheet = workbook.getSheetAt(0);
        if (sheet.getPhysicalNumberOfRows() < 2)
            throw new RuntimeException(); //TODO нормальные исключения
        if (sheet.getRow(0).getPhysicalNumberOfCells() == 0)
            throw new RuntimeException(); //TODO нормальные исключения
    }

    @Override
    public void close() throws IOException {
        source.close();
        cutSource.close();
        result.close();
    }
}
