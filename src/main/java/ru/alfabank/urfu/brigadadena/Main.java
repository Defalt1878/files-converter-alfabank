package ru.alfabank.urfu.brigadadena;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.alfabank.urfu.brigadadena.excel.converter.ExcelConverter;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        final String sourcePath = "src/main/resources/source.xlsx";
        final String samplePath = "src/main/resources/sample.xlsx";
        final String resultPath = "src/main/resources/result.xlsx";

        try (
            var source = new FileInputStream(sourcePath);
            var sample = new FileInputStream(samplePath)
        ) {
            var sourceWB = new XSSFWorkbook(source);
            var sampleWB = new XSSFWorkbook(sample);
            var converter = new ExcelConverter(sourceWB, sampleWB);
            converter.divideColumns(0, " ");
            converter.connectColumns(0, 0);
            converter.connectColumns(1, 1);
            converter.connectColumns(2, 2);
            converter.connectColumns(3, 3);
            converter.cancelLast();

            var result = converter.getFinalResult();
            try (var outputStream = new FileOutputStream(resultPath)) {
                result.write(outputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
