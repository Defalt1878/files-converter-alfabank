package ru.alfabank.urfu.brigadadena.excel.sheetConnector;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class SheetConnector {
    private List<ColumnsConnector> columnsConnectors = new ArrayList<>();

    public SheetConnector() {

    }

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

    public boolean removeSrcIntersections(int[] srsColumnsNums) {
        var prevSize = columnsConnectors.size();
        var numsSet = Arrays.stream(srsColumnsNums).boxed().collect(Collectors.toSet());
        columnsConnectors = columnsConnectors.stream()
            .filter(connector -> !numsSet.contains(connector.getSrcColumnNum()))
            .collect(Collectors.toList());
        return prevSize != columnsConnectors.size();
    }
}
