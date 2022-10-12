package ru.alfabank.urfu.brigadadena;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.alfabank.urfu.brigadadena.excel.columnsmodifiets.ColumnsDateCombiner;
import ru.alfabank.urfu.brigadadena.excel.columnsmodifiets.ColumnsDateDivider;
import ru.alfabank.urfu.brigadadena.excel.columnsmodifiets.ColumnsShifter;
import ru.alfabank.urfu.brigadadena.excel.columnsmodifiets.ColumnsSimpleDivider;
import ru.alfabank.urfu.brigadadena.excel.converter.ExcelConverter;

import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        final String sourcePath = "src/main/resources/source.xlsx";
        final String samplePath = "src/main/resources/sample.xlsx";
        final String resultPath = "src/main/resources/result.xlsx";

        try (
            var converter = new ExcelConverter(new XSSFWorkbook(sourcePath), new XSSFWorkbook(samplePath));
            var outputStream = new FileOutputStream(resultPath)
        ) {
            converter.addModifier(new ColumnsSimpleDivider(0, new int[]{0, 1, 2}, " "));
            converter.addModifier(new ColumnsShifter(1, 5));
            converter.addModifier(new ColumnsShifter(2, 6));
            converter.addModifier(new ColumnsShifter(3, 4));
            converter.addModifier(new ColumnsShifter(4, 8));
            converter.addModifier(new ColumnsShifter(5, 7));
            converter.addModifier(new ColumnsShifter(6, 9));
            converter.addModifier(new ColumnsShifter(7, 13));
            converter.addModifier(new ColumnsShifter(8, 12));
            converter.addModifier(new ColumnsDateCombiner(new int[]{10, 11}, 10));
            converter.addModifier(new ColumnsDateCombiner(new int[]{12, 13}, 11));
            converter.addModifier(new ColumnsDateDivider(14, new int[]{14, 15}));
            var result = converter.getResult();
            result.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}