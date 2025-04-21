package com.divakar.bankToBudget.fileProcessing.extractors;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class ExcelColumnExtractor {

    public static List<String> extractTransactions(InputStream is, String filename) throws Exception {
        return extractColumnValues(is, filename, 1, false);
    }

    public static List<String> extractAmounts(InputStream is, String filename) throws Exception {
        return extractColumnValues(is, filename, 3, true);
    }

    private static Workbook getWorkbook(InputStream is, String filename) throws Exception {
        if (filename.endsWith(".xlsx")) return new XSSFWorkbook(is);
        if (filename.endsWith(".xls")) return new HSSFWorkbook(is);
        throw new IllegalArgumentException("Unsupported file type: " + filename);
    }

    private static List<String> extractColumnValues(InputStream is, String filename, int columnIndex, boolean parseAsDouble) throws Exception{
        List<String> values = new ArrayList<>();

        try (Workbook workbook = getWorkbook(is, filename)) {
            Sheet sheet = workbook.getSheetAt(0);
            int rowIndex = 0;

            for (Row row : sheet) {
                rowIndex++;
                if (rowIndex < 17) continue;

                Cell cell = row.getCell(columnIndex);
                if (cell == null) continue;

                String cellValue = cell.toString().trim();
                if (parseAsDouble) {
                    try {
                        double amount = cellValue.isBlank() ? 0 :
                                NumberFormat.getNumberInstance(new Locale("en", "IN")).parse(cellValue).doubleValue();
                        values.add(String.valueOf(amount));
                    } catch (Exception e) {
                        System.out.println("Failed to parse amount: '" + cellValue + "' | Error: " + e.getMessage());
                        values.add("Number format exception");
                    }
                } else {
                    values.add(cellValue);
                }
            }
        }

        return values;
    }
}
