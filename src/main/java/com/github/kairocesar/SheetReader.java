package com.github.kairocesar;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Objects;

public class SheetReader {
    public static void readSheetsAndSearchValues(String folderPath) {

        // Nomes das colunas que queremos verificar
        String[] columnsToCheck = {"PIS", "COFINS", "INSS", "IRRF", "CSLL"};

        File folder = new File(folderPath);
        File[] listOfFiles = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".xlsx"));

        if (listOfFiles != null) {
            for (File file : listOfFiles) {
                if (file.isFile()) {
                    try (FileInputStream fis = new FileInputStream(file);
                         Workbook workbook = new XSSFWorkbook(fis)) {

                        Sheet sheet = workbook.getSheet("Serviços Tomados");
                        if (sheet != null) {
                            for (String columnName : columnsToCheck) {
                                int columnIndex = findColumnIndex(sheet, columnName);
                                if (columnIndex != -1 && hasFilledValues(sheet, columnIndex)) {
                                    System.out.println("A planilha " + file.getName() + " contém valores de " + columnName
                                            + " retidos. Favor verificar!");
                                }
                            }
                        }
                    } catch (IOException e) {
                        System.out.println("Erro ao ler a planilha " + file.getName() + ": " + e.getMessage());
                    }
                }
            }
        }
    }

    private static int findColumnIndex(Sheet sheet, String columnName) {
        Row firstRow = sheet.getRow(0);
        if (firstRow != null) {
            for (Cell cell : firstRow) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1; // Coluna não encontrada
    }

    private static boolean hasFilledValues(Sheet sheet, int columnIndex) {
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Ignora a primeira linha (cabeçalho)
            Cell cell = row.getCell(columnIndex);
            if (isNotBlank(cell) && isString(cell) && (!cell.getStringCellValue().contains("0") && !cell.getStringCellValue().isEmpty())) {
                return true;
            } else if (isNotBlank(cell) && isNumeric(cell) && cell.getNumericCellValue() > 0) {
                return true;
            }
        }
        return false;
    }

    private static boolean isString(Cell cell) {
        try {
            cell.getStringCellValue();
            return true;
        } catch (IllegalStateException e) {
            return false;
        }
    }

    private static boolean isNumeric(Cell cell) {
        try {
            cell.getNumericCellValue();
            return true;
        } catch (IllegalStateException e) {
            return false;
        }
    }

    private static boolean isNotBlank(Cell cell) {
        return !Objects.isNull(cell) && cell.getCellType() != CellType.BLANK;
    }
}
