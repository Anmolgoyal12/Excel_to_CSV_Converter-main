package org.csdconverter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;

import java.io.*;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Logger;
import java.util.stream.Collectors;

/**
 * This class converts data from Excel sheets to CSV files based on configuration.
 */
public class MainCSD {
    // Output folder path
    private static final String BASE_OUTPUT_DIR = "D:/Excel_to_CSV_Converter-main/BASE_OUTPUT_DIRECTORY";
    private static final Logger logger = Logger.getLogger(MainCSD.class.getName());

    /**
     * Main method to initiate the Excel to CSV conversion process.
     *
     * @param args Command-line arguments (not used in this application)
     */
    public static void main(String[] args) {
        // Paths to configuration and Excel files
        String configFilePath = "D:/CSV Automation/CSD_TO_CSV.xlsx";
        String excelFilePath = "D:/CSV Automation/CSD - Internal.xlsx";

        // Load sheet configurations from the config file
        List<SheetConfig> sheetConfigs = loadSheetConfigs(configFilePath);

        // Process each sheet based on its configuration
        for (SheetConfig config : sheetConfigs) {
            String sheetName = config.getSheetName();
            String csvFilePath = Paths.get(BASE_OUTPUT_DIR, config.getOutputDirectory(), config.getCsvName()).toString();
            try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath))) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet != null) {
                    // Extract data from the sheet based on configuration
                    List<List<String>> extractedData = extractDataFromSheet(sheet, config);

                    // Apply transposition if required
                    if (config.isTranspose() && !config.getExcludeFromTranspose().contains(sheetName)) {
                        extractedData = transposeData(extractedData);
                    }

                    // Apply advance condition to headers
                    applyAdvanceConditionToHeaders(extractedData);
                    // Clean up headers (remove asterisks)
                    cleanUpHeaders(extractedData);
                    // Write extracted data to CSV file
                    writeCSV(csvFilePath, extractedData);
                } else {
                    throw new Exception("Sheet not found: " + sheetName);
                }
            } catch (Exception e) {
                logger.severe("Error processing sheet: " + sheetName + ". " + e.getMessage());
            }
        }

        logger.info("Conversion completed successfully.");
    }

    /**
     * Loads sheet configurations from the specified Excel file.
     *
     * @param configFilePath Path to the configuration Excel file
     * @return List of SheetConfig objects representing each sheet's configuration
     */
    private static List<SheetConfig> loadSheetConfigs(String configFilePath) {
        List<SheetConfig> sheetConfigs = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(configFilePath))) {
            Sheet configSheet = workbook.getSheetAt(0);
            for (Row row : configSheet) {
                if (row.getRowNum() == 0) continue; // Skip header row
                SheetConfig config = new SheetConfig(
                        getCellValue(row.getCell(1)), // Sheet name
                        getCellValue(row.getCell(2)), // CSV name
                        getTextBooleanCellValue(row.getCell(3)), // Is Transpose
                        getTextBooleanCellValue(row.getCell(4)), // Is Comment Read
                        getCellValue(row.getCell(5)), // Range for specified column
                        getStringListCellValue(row.getCell(6)), // Exclude from Transpose list
                        getCellValue(row.getCell(7)) // Output Directory
                );
                sheetConfigs.add(config);
            }
        } catch (IOException e) {
            logger.severe("Error loading sheet configurations: " + e.getMessage());
        }
        return sheetConfigs;
    }

    /**
     * Standardizes the header by converting to lowercase, replacing spaces with underscores,
     * and removing trailing underscores.
     * Specifically handles converting 'username' to 'user_name'.
     *
     * @param input Input string to be standardized
     * @return Standardized header
     */
    public static String standardizeHeader(String input) {
        if (input == null || input.isEmpty()) {
            return input;
        }

        // Handle asterisks
        input = input.replace("*", "");

        // Remove trailing underscore
        while (input.endsWith("_")) {
            input = input.substring(0, input.length() - 1);
        }

        // Handle specific cases
        String lowerCaseInput = input.toLowerCase();
        if ("username".equalsIgnoreCase(lowerCaseInput)) {
            return "user_name";
        }
        //Handle for specific case Input
        String list=input.toLowerCase();
        if("primarykeylist".equalsIgnoreCase(list))
        {
            return "primarykey_list";
        }

        return lowerCaseInput.replaceAll("\\s+", "_");
    }

    /**
     * Utility method to get cell value as string.
     *
     * @param cell Excel Cell object from which to retrieve the value
     * @return String value of the cell
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue()); // Convert to integer

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    /**
     * Utility method to get boolean value from text representation.
     *
     * @param cell Excel Cell object containing text representation of boolean
     * @return Boolean value parsed from the text representation
     */
    private static boolean getTextBooleanCellValue(Cell cell) {
        if (cell == null) {
            return false;
        }

        if (cell.getCellType() == CellType.BOOLEAN) {
            return cell.getBooleanCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            String cellValue = cell.getStringCellValue().trim();

            return "true".equalsIgnoreCase(cellValue);
        }

        return false;
    }

    /**
     * Utility method to get list of strings from a cell containing comma-separated values.
     *
     * @param cell Excel Cell object containing comma-separated values
     * @return List of strings parsed from the cell value
     */
    private static List<String> getStringListCellValue(Cell cell) {
        List<String> stringList = new ArrayList<>();
        if (cell != null && cell.getCellType() == CellType.STRING) {
            String[] values = cell.getStringCellValue().split(",");
            for (String value : values) {
                stringList.add(value.trim());
            }
        }
        return stringList;
    }

    /**
     * Extracts data from the specified sheet based on configuration, including specified ranges.
     *
     * @param sheet  Excel Sheet object from which to extract data
     * @param config Configuration for how to extract data from the sheet
     * @return List of lists containing extracted data
     */
    private static List<List<String>> extractDataFromSheet(Sheet sheet, SheetConfig config) {
        List<List<String>> data = new ArrayList<>();

        boolean shouldTranspose = config.isTranspose();
        logger.info("Sheet: " + sheet.getSheetName() + " - Should Transpose: " + shouldTranspose);
        int startColumn = 1;

        Row headerRow = sheet.getRow(0);
        int commentColumnIndex = -1;

        if (headerRow != null) {
            for (Cell cell : headerRow) {
                if ("Comment".equals(cell.getStringCellValue()) || "Comments".equals(cell.getStringCellValue())) {
                    commentColumnIndex = cell.getColumnIndex();
                    break;
                }
            }
        }

        String range = config.getRange();
        int startRow = shouldTranspose ? 2 : 0;

        List<Integer> rowIndices = new ArrayList<>();

        if (range != null && !range.isEmpty() && !"NA".equalsIgnoreCase(range)) {
            String[] parts = range.split(",");
            for (String part : parts) {
                if (part.contains("-")) {
                    String[] bounds = part.split("-");
                    try {
                        int start = Integer.parseInt(bounds[0].trim());
                        int end = Integer.parseInt(bounds[1].trim());
                        for (int i = start; i <= end; i++) {
                            rowIndices.add(i - 1); // Adjust to 0-based index
                        }
                    } catch (NumberFormatException e) {
                        logger.severe("Invalid range format: " + range);
                    }
                } else if (part.matches("\\d+")) {
                    int row = Integer.parseInt(part.trim()) - 1; // Adjust to 0-based index
                    rowIndices.add(row);
                } else if (part.matches("[A-Z]+\\d+")) {
                    CellReference cellReference = new CellReference(part.trim());
                    int row = cellReference.getRow();
                    rowIndices.add(row);
                } else {
                    logger.severe("Invalid range format: " + range);
                }
            }
        }

        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            if (!rowIndices.isEmpty() && !rowIndices.contains(i)) continue;

            Row row = sheet.getRow(i);
            if (row == null) continue;

            if (config.isCommentRead() != null && config.isCommentRead()) {
                Cell firstCell = row.getCell(0);
                String cellValue = getCellValue(firstCell);
                if (cellValue != null && cellValue.startsWith("#")) {
                    continue;
                }
            }
            List<String> rowData = new ArrayList<>();
            for (int j = startColumn; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (!config.isCommentRead() && j == commentColumnIndex) {
                    continue;
                }
                rowData.add(getCellValue(cell));
            }

            data.add(rowData);
        }

        return data;
    }
    /**
     * Transposes the data.
     *
     * @param data List of lists containing data to be transposed
     * @return List of lists containing transposed data
     */
    private static List<List<String>> transposeData(List<List<String>> data) {
        List<List<String>> transposedData = new ArrayList<>();
        if (data.isEmpty() || data.get(0).isEmpty()) return transposedData;

        int colCount = data.get(0).size();

        for (int col = 0; col < colCount; col++) {
            List<String> transposedRow = new ArrayList<>();
            for (List<String> currentRow : data) {
                if (col < currentRow.size()) {
                    transposedRow.add(currentRow.get(col));
                } else {
                    transposedRow.add(""); // or handle missing value as per your requirement
                }
            }
            transposedData.add(transposedRow);
        }

        return transposedData;
    }

    /**
     * Applies advance condition to headers by converting them to lowercase and replacing spaces with underscores.
     *
     * @param data List of lists containing data with headers to be standardized
     */
    private static void applyAdvanceConditionToHeaders(List<List<String>> data) {
        if (data == null || data.isEmpty()) {
            return;
        }

        List<String> headers = data.get(0);
        headers.replaceAll(MainCSD::standardizeHeader);
    }

    /**
     * Cleans up the headers by removing asterisks from each cell.
     *
     * @param data List of lists containing data to be cleaned up
     */
    private static void cleanUpHeaders(List<List<String>> data) {
        if (data == null || data.isEmpty()) {
            return;
        }

        List<String> headers = data.get(0);
        headers.replaceAll(s -> s.replace("*", ""));
    }

    /**
     * Escapes CSV data by handling commas, newlines, and double quotes.
     *
     * @param data Data to be escaped
     * @return Escaped CSV data
     */
    private static String escapeCsvData(String data) {
        if (data.contains(",") || data.contains("\n") || data.contains("\"")) {
            data = data.replace("\"", "\"\"");
            data = "\"" + data + "\"";
        }
        return data;
    }

    /**
     * Writes data to a CSV file.
     *
     * @param csvFilePath Path to the output CSV file
     * @param data        Data to be written to the CSV file
     */
    private static void writeCSV(String csvFilePath, List<List<String>> data) {
        File outputFile = new File(csvFilePath);
        if (!outputFile.getParentFile().exists() && !outputFile.getParentFile().mkdirs()) {
            logger.severe("Failed to create output directories for: " + csvFilePath);
            return;
        }
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(outputFile))) {
            for (List<String> row : data) {
                String line = row.stream().map(MainCSD::escapeCsvData).collect(Collectors.joining(","));
                writer.write(line);
                writer.newLine();
            }
        } catch (IOException e) {
            logger.severe("Error writing CSV file: " + csvFilePath + ". " + e.getMessage());
        }
    }
}
