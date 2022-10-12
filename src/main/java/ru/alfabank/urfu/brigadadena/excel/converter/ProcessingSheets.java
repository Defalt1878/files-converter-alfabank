package ru.alfabank.urfu.brigadadena.excel.converter;

import org.apache.poi.ss.usermodel.Sheet;

public record ProcessingSheets(Sheet source, Sheet sample, Sheet result) {

}
