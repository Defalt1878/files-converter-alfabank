package ru.alfabank.urfu.brigadadena;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.alfabank.urfu.brigadadena.excel.converter.ExcelConverter;
import ru.alfabank.urfu.brigadadena.excel.util.ExcelHelper;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        final String sourcePath = "src/main/resources/source.xlsx";
        final String samplePath = "src/main/resources/sample.xlsx";
        final String resultPath = "src/main/resources/result.xlsx";

        var scanner = new Scanner(System.in);

        try (
            var source = new FileInputStream(sourcePath);
            var sample = new FileInputStream(samplePath)
        ) {
            var sourceWB = new XSSFWorkbook(source);
            var sampleWB = new XSSFWorkbook(sample);
            var converter = new ExcelConverter(sourceWB, sampleWB);
            while (true) {
                outputResult(converter);
                System.out.print("Input command: ");
                var cmd = scanner.nextLine();
                if (executeCommand(cmd, converter))
                    break;
            }

            var result = converter.getFinalResult();
            try (var outputStream = new FileOutputStream(resultPath)) {
                result.write(outputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static boolean executeCommand(String command, ExcelConverter converter) throws IOException {
        var parts = command.split(" ");
        switch (parts[0]) {
            case "combine" -> {
                var splitter = command.split("\"")[1];
                var columnsNums = Arrays.stream(parts)
                    .filter(str -> str.matches("\\d+"))
                    .mapToInt(Integer::parseInt)
                    .map(num -> num - 1)
                    .toArray();
                converter.combineColumns(columnsNums, splitter);
            }
            case "divide" -> {
                var splitter = command.split("\"")[1];
                var columnNum = Integer.parseInt(parts[1]) - 1;
                converter.divideColumns(columnNum, splitter);
            }
            case "connect" -> {
                var sourceNum = Integer.parseInt(parts[1]) - 1;
                var resultNum = Integer.parseInt(parts[2]) - 1;
                converter.connectColumns(sourceNum, resultNum);
            }
            case "cancel" -> {
                converter.cancelLast();
            }
            case "convert" -> {
                return true;
            }
            default -> {
                System.out.println("Command error!");
                return false;
            }
        }
        return false;
    }

    private static void outputResult(ExcelConverter converter) {
        var source = converter.getSourceExample();
        var result = converter.getResultExample();
        System.out.println();
        printTable(ExcelHelper.toStringMatrix(source.getSheetAt(0)));
        System.out.println();
        printTable(ExcelHelper.toStringMatrix(result.getSheetAt(0)));
        System.out.println();
    }

    private static void printTable(String[][] table) {
        for (var line : table) {
            for (var value : line)
                System.out.printf("%-30s", value);
            System.out.println();
        }
    }
}
