package com.divakar.bankToBudget.fileProcessing.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFPieChartData;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static com.divakar.bankToBudget.fileProcessing.extractors.ExcelColumnExtractor.extractAmounts;
import static com.divakar.bankToBudget.fileProcessing.extractors.ExcelColumnExtractor.extractTransactions;
import static com.divakar.bankToBudget.fileProcessing.extractors.NotesExtractor.getNotesFromTransactionsRorTMB;


@Service
public class FileProcessingService {

    public String processExcelFile(MultipartFile file) throws Exception {
        System.out.println("Processing Excel file.");
        final List<String> columnBData = extractTransactions(file.getInputStream(), file.getOriginalFilename());
        final List<String> notes = getNotesFromTransactionsRorTMB(columnBData);
        final List<String> amounts = extractAmounts(file.getInputStream(), file.getOriginalFilename());

        Path tempFile = Files.createTempFile("output", ".xlsx");
        String outputFilePath = tempFile.toString();

        writeNotesAndAmountsToExcel(notes, amounts, outputFilePath);
        return outputFilePath;
    }

    public static void writeNotesAndAmountsToExcel(List<String> notes, List<String> amounts, String outputFilePath) throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Output");

            // Header style
            XSSFCellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Header row
            Row headerRow = sheet.createRow(0);
            Cell categoryHeader = headerRow.createCell(0);
            Cell amountHeader = headerRow.createCell(1);

            categoryHeader.setCellValue("Category");
            categoryHeader.setCellStyle(headerStyle);

            amountHeader.setCellValue("Amount");
            amountHeader.setCellStyle(headerStyle);

            // Fill data
            int rowCount = Math.min(notes.size(), amounts.size());
            double totalAmount = 0;

            // Process data and calculate total
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(notes.get(i));

                Cell amountCell = row.createCell(1);
                try {
                    String cleanAmount = amounts.get(i).replaceAll("[^0-9.-]", "");
                    double amountValue = Double.parseDouble(cleanAmount);
                    amountCell.setCellValue(amountValue);
                    totalAmount += amountValue;
                } catch (NumberFormatException e) {
                    amountCell.setCellValue(amounts.get(i));
                }
            }

            // Auto-size columns
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);

            // Group small categories together
            XSSFSheet consolidatedSheet = workbook.createSheet("Consolidated");

            // Create consolidated data structure to combine small values
            Map<String, Double> consolidatedData = new HashMap<>();
            final double THRESHOLD = totalAmount * 0.02; // 2% threshold

            // Header row for consolidated sheet
            Row consolidatedHeader = consolidatedSheet.createRow(0);
            Cell consolidatedCategoryHeader = consolidatedHeader.createCell(0);
            Cell consolidatedAmountHeader = consolidatedHeader.createCell(1);
            Cell consolidatedPercentHeader = consolidatedHeader.createCell(2);

            consolidatedCategoryHeader.setCellValue("Category");
            consolidatedCategoryHeader.setCellStyle(headerStyle);
            consolidatedAmountHeader.setCellValue("Amount");
            consolidatedAmountHeader.setCellStyle(headerStyle);
            consolidatedPercentHeader.setCellValue("Percentage");
            consolidatedPercentHeader.setCellStyle(headerStyle);

            // Process data for consolidation
            for (int i = 0; i < rowCount; i++) {
                String category = notes.get(i);
                double amount = 0;

                try {
                    String cleanAmount = amounts.get(i).replaceAll("[^0-9.-]", "");
                    amount = Double.parseDouble(cleanAmount);
                } catch (NumberFormatException e) {
                    continue;
                }

                // Group by category and sum amounts
                if (consolidatedData.containsKey(category)) {
                    consolidatedData.put(category, consolidatedData.get(category) + amount);
                } else {
                    consolidatedData.put(category, amount);
                }
            }

            // Separate small categories into "Other"
            Map<String, Double> chartData = new LinkedHashMap<>();
            double otherTotal = 0;

            // Sort categories by amount in descending order
            List<Map.Entry<String, Double>> sortedData = new ArrayList<>(consolidatedData.entrySet());
            sortedData.sort((a, b) -> b.getValue().compareTo(a.getValue()));

            // Add top categories and group small ones
            int chartRowIndex = 1;
            for (Map.Entry<String, Double> entry : sortedData) {
                String category = entry.getKey();
                double amount = entry.getValue();

                // Add to chart data or "Other" based on threshold
                if (amount >= THRESHOLD && chartRowIndex <= 10) { // Limit to 10 slices
                    chartData.put(category, amount);
                } else {
                    otherTotal += amount;
                }

                // Add to consolidated sheet
                Row dataRow = consolidatedSheet.createRow(chartRowIndex++);
                dataRow.createCell(0).setCellValue(category);
                dataRow.createCell(1).setCellValue(amount);
                dataRow.createCell(2).setCellValue(amount / totalAmount);
            }

            // Add "Other" category if needed
            if (otherTotal > 0) {
                chartData.put("Other", otherTotal);
            }

            // Format percentage column
            XSSFCellStyle percentStyle = workbook.createCellStyle();
            percentStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));

            for (int i = 1; i < chartRowIndex; i++) {
                Cell percentCell = consolidatedSheet.getRow(i).getCell(2);
                percentCell.setCellStyle(percentStyle);
            }

            // Auto-size columns
            consolidatedSheet.autoSizeColumn(0);
            consolidatedSheet.autoSizeColumn(1);
            consolidatedSheet.autoSizeColumn(2);

            // Create a new sheet for the chart
            XSSFSheet chartSheet = workbook.createSheet("Chart");

            // Add data for the chart
            Row chartHeaderRow = chartSheet.createRow(0);
            chartHeaderRow.createCell(0).setCellValue("Category");
            chartHeaderRow.createCell(1).setCellValue("Amount");
            chartHeaderRow.createCell(2).setCellValue("Percentage");

            chartHeaderRow.getCell(0).setCellStyle(headerStyle);
            chartHeaderRow.getCell(1).setCellStyle(headerStyle);
            chartHeaderRow.getCell(2).setCellStyle(headerStyle);

            // Add data rows
            int rowIdx = 1;
            for (Map.Entry<String, Double> entry : chartData.entrySet()) {
                Row dataRow = chartSheet.createRow(rowIdx++);
                dataRow.createCell(0).setCellValue(entry.getKey());
                dataRow.createCell(1).setCellValue(entry.getValue());
                dataRow.createCell(2).setCellValue(entry.getValue() / totalAmount);
                dataRow.getCell(2).setCellStyle(percentStyle);
            }

            // Auto-size columns
            chartSheet.autoSizeColumn(0);
            chartSheet.autoSizeColumn(1);
            chartSheet.autoSizeColumn(2);

            // Create drawing for the chart sheet
            XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 1, 15, 20);

            // Create chart
            XSSFChart chart = drawing.createChart(anchor);
            chart.setTitleText("Expense Categories");
            chart.setTitleOverlay(false);

            // Define chart legend
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);

            // Define data ranges for the chart
            int chartDataSize = chartData.size();
            XDDFDataSource<String> categories = XDDFDataSourcesFactory.fromStringCellRange(chartSheet,
                    new CellRangeAddress(1, chartDataSize, 0, 0));
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(chartSheet,
                    new CellRangeAddress(1, chartDataSize, 1, 1));

            // Create chart data
            XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
            XDDFChartData.Series series = data.addSeries(categories, values);

            // Remove the "Series1" prefix from the series name
            series.setTitle("", null);

            // Make sure we have varying colors for pie slices
            if (data instanceof XDDFPieChartData) {
                ((XDDFPieChartData) data).setVaryColors(true);
            }

            // Plot the chart
            chart.plot(data);

            // Advanced XML manipulation to fix labels
            try {
                // Get the underlying XML representation
                org.openxmlformats.schemas.drawingml.x2006.chart.CTChart ctChart = chart.getCTChart();
                org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea plotArea = ctChart.getPlotArea();

                if (plotArea.getPieChartArray().length > 0) {
                    org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart pieChart = plotArea.getPieChartArray(0);

                    // Set up data labels for each series
                    for (org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer ser : pieChart.getSerArray()) {
                        // Remove series name if present
                        if (ser.isSetTx()) {
                            ser.unsetTx();
                        }

                        // Configure data labels
                        org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls dLbls = ser.isSetDLbls() ?
                                ser.getDLbls() : ser.addNewDLbls();

                        // Configure what's shown in labels
                        dLbls.addNewShowVal().setVal(false);
                        dLbls.addNewShowPercent().setVal(true);
                        dLbls.addNewShowCatName().setVal(true);
                        dLbls.addNewShowLegendKey().setVal(false);

                        // Critical - ensure series name is not shown
                        if (!dLbls.isSetShowSerName()) {
                            dLbls.addNewShowSerName().setVal(false);
                        } else {
                            dLbls.getShowSerName().setVal(false);
                        }

                        // Position the labels outside
                        org.openxmlformats.schemas.drawingml.x2006.chart.CTDLblPos pos = dLbls.isSetDLblPos() ?
                                dLbls.getDLblPos() : dLbls.addNewDLblPos();
                        pos.setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.OUT_END);

                        // Note: No separator setting as it's not available in your POI version
                    }
                }
            } catch (Exception e) {
                System.err.println("Could not fully configure chart labels: " + e.getMessage());

                // Fallback to the simpler method if advanced configuration fails
                try {
                    org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea plotArea = chart.getCTChart().getPlotArea();
                    if (plotArea.getPieChartList().size() > 0) {
                        org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart pieChart = plotArea.getPieChartArray(0);
                        if (pieChart.getSerList().size() > 0) {
                            org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls dLbls = pieChart.getSerArray(0).addNewDLbls();
                            dLbls.addNewShowVal().setVal(false);
                            dLbls.addNewShowPercent().setVal(true);
                            dLbls.addNewShowCatName().setVal(true);
                            dLbls.addNewShowLegendKey().setVal(false);
                            dLbls.addNewShowSerName().setVal(false);  // Turn off series name
                        }
                    }
                } catch (Exception ex) {
                    System.err.println("Could not configure data labels: " + ex.getMessage());
                }
            }

            // Set chart sheet as active
            workbook.setActiveSheet(workbook.getSheetIndex("Chart"));

            // Write to file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                workbook.write(fos);
            }
        }
    }
}