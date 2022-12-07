package ru.alfabank.urfu.brigadadena.excel.sheetConnector;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class SheetConnector {
    private List<ColumnsConnector> columnsConnectors = new ArrayList<>();

    public void newConnector(ColumnsConnector connector) {
        var intersection = columnsConnectors.stream()
            .filter(conn -> connector.getDstColumnNum() == conn.getDstColumnNum())
            .findFirst();
        intersection.ifPresent(columnsConnectors::remove);
        columnsConnectors.add(connector);
    }

    public void applyAll(Sheet src, Sheet dst, int rowsCount) {
        columnsConnectors.forEach(connector -> connector.apply(src, dst, rowsCount));
    }

    public void applyLast(Sheet src, Sheet dst, int rowsCount) {
        columnsConnectors.get(columnsConnectors.size() - 1).apply(src, dst, rowsCount);
    }

    public void updateConnectors(int[] srsColumnsNums, int[] resultColumnNums) {
        var srcNumsSet = Arrays.stream(srsColumnsNums).boxed().collect(Collectors.toSet());
        var resultNumsSet = Arrays.stream(resultColumnNums).boxed().collect(Collectors.toSet());

        var updatedConnectors = new ArrayList<ColumnsConnector>();
        for (var columnsConnector : columnsConnectors) {
            var srcColumnNum = columnsConnector.getSrcColumnNum();
            if (srcNumsSet.contains(srcColumnNum))
                continue;

            var srcNumsBefore = srcNumsSet.stream().filter(num -> num <= srcColumnNum).count();
            var resultNumsBefore = resultNumsSet.stream().filter(num -> num <= srcColumnNum).count();
            var difference = (int) (resultNumsBefore - srcNumsBefore);
            if (difference == 0)
                updatedConnectors.add(columnsConnector);
            else if (difference > 0) {
                updatedConnectors.add(new ColumnsConnector(
                    srcColumnNum + resultNumsSet.size() - srcNumsSet.size(),
                    columnsConnector.getDstColumnNum(),
                    columnsConnector.getDstStyle()
                ));
            } else
                updatedConnectors.add(new ColumnsConnector(
                    srcColumnNum + difference,
                    columnsConnector.getDstColumnNum(),
                    columnsConnector.getDstStyle()
                ));
        }
        this.columnsConnectors = updatedConnectors;
    }

    public boolean removeFor(int resultColumnNum) {
        var toRemove = columnsConnectors.stream()
            .filter(connector -> connector.getDstColumnNum() == resultColumnNum)
            .findFirst();
        if (toRemove.isEmpty())
            return false;
        columnsConnectors.remove(toRemove.get());
        return true;
    }
}
