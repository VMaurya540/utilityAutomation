package com.example.utilityAutomation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;


@Service
public class ExcelReaderService {

    //public static void main(String[] args) {

        // Read the DB Objects sheet and generate comments
        public void readDBObjectsSheet() {
            // Increase the POI memory limit
            IOUtils.setByteArrayMaxOverride(300_000_000); // 300MB limit

            // Define File Paths
            String filePath = "C:\\Users\\Vishal Kumar\\Desktop\\Master_Tracker\\Boral_Master_tracker_v3_detailed.xlsx";
            String sheetName = "DB Objects";
            String outputPath = "C:\\Users\\Vishal Kumar\\Desktop\\utility_comments.txt";

            // Check File Exists
            File file = new File(filePath);
            if (!file.exists()) {
                System.err.println(" ERROR: Excel file not found at: " + file.getAbsolutePath());
                return;
            }

            try {
                // Open Excel File
                FileInputStream fileInputStream = new FileInputStream(file);
                Workbook workbook = new XSSFWorkbook(fileInputStream);
                Sheet sheet = workbook.getSheet(sheetName);

                if (sheet == null) {
                    System.err.println(" ERROR: Excel sheet '" + sheetName + "' not found!");
                    workbook.close();
                    return;
                }

                // Find Column Indexes
                Map<String, Integer> columnIndexes = getColumnIndexes(sheet);
                if (columnIndexes.isEmpty()) {
                    System.err.println(" ERROR: Required column headers not found!");
                    workbook.close();
                    return;
                }

                // Generate Comments
                List<String> comments = generateComments(sheet, columnIndexes);
                workbook.close();

                // Save Comments to File
                saveCommentsToFile(comments, outputPath);
                System.out.println("Comments saved to: " + outputPath);

            } catch (IOException e) {
                e.printStackTrace();
            }
        }


    // Identify Column Indexes Dynamically
    private static Map<String, Integer> getColumnIndexes(Sheet sheet) {
        Map<String, Integer> columnIndexes = new HashMap<>();

        Row headerRow = sheet.getRow(1);
        if (headerRow == null) {
            System.err.println("ERROR: Header row is empty!");
            return columnIndexes;
        }

        System.out.println("🔍 Headers found in Row :");
        for (Cell cell : headerRow) {
            // String header = cell.getStringCellValue().trim();
            String header = getCellValue(headerRow, cell.getColumnIndex());
            System.out.println(" - " + header);


            switch (header) {
                case "OBJECT_NAME":
                case "REFERENCED_TYPE":
                case "OBJECT_SEARCH_NAME":
                case "CHANGE_TYPE":
                case "ADDITIONAL_INFO1_1":
                    columnIndexes.put(header, cell.getColumnIndex());
                    break;
            }
        }

        return columnIndexes;
    }

    // Generate Comments Based on Utility Report
    private static List<String> generateComments(Sheet sheet, Map<String, Integer> columnIndexes) {

        List<String> comments = new ArrayList<>();
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
        String currentDate = sdf.format(new Date());

        // Map to store consolidated messages per object
        Map<String, Map<String, Set<String>>> consolidatedComments = new LinkedHashMap<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Skip header row

            // Read values
            String objectName = getCellValue(row, columnIndexes.get("OBJECT_NAME"));
            String referencedType = getCellValue(row, columnIndexes.get("REFERENCED_TYPE"));
            String objectSearchName = getCellValue(row, columnIndexes.get("OBJECT_SEARCH_NAME"));
            String changeType = getCellValue(row, columnIndexes.get("CHANGE_TYPE"));
            String additionalInfo = getCellValue(row, columnIndexes.get("ADDITIONAL_INFO1_1"));

            // Log values to check if data is being read correctly
            System.out.println("Reading row " + row.getRowNum() + ": " + objectName + ", " + referencedType + ", " + objectSearchName);

            if (objectSearchName.isEmpty() || changeType.isEmpty()) continue; // Skip empty rows

            // Initialize object if not present
            consolidatedComments.putIfAbsent(objectName, new LinkedHashMap<>());

            // Identify change category
//            String keyCategory = "";
//            if (referencedType.equalsIgnoreCase("PACKAGE")) {
//                keyCategory = changeType.contains("ADDED") ? "proceduresAdded" : "functionsRemoved";
//            } else if (referencedType.equalsIgnoreCase("TABLE")) {
//                keyCategory = changeType.contains("ADDED") ? "columnsAdded" : "columnsModified";
//            } else if (referencedType.equalsIgnoreCase("VIEW")) {
//                keyCategory = "viewsModified";
//            }
            // Identify change category
            String keyCategory = "";
            if (referencedType.equalsIgnoreCase("PACKAGE")) {
                if (changeType.equalsIgnoreCase("PROC_MODIFIED")) {
                    keyCategory = "proceduresModified";
                } else {
                    keyCategory = changeType.contains("ADDED") ? "proceduresAdded" : "functionsRemoved";
                }
            } else if (referencedType.equalsIgnoreCase("TABLE")) {
                keyCategory = "columnsModified";
            } else if (referencedType.equalsIgnoreCase("VIEW")) {
                keyCategory = "viewsModified";
            }

            // Store values in the respective category
            consolidatedComments.get(objectName).putIfAbsent(keyCategory, new LinkedHashSet<>());
            consolidatedComments.get(objectName).get(keyCategory).add(objectSearchName + (additionalInfo.isEmpty() ? "" : " (" + additionalInfo + ")"));
        }

        // **Generating final consolidated comments with object name**
        List<String> finalComments = new ArrayList<>();
        for (Map.Entry<String, Map<String, Set<String>>> entry : consolidatedComments.entrySet()) {
            String objectName = entry.getKey();
            StringBuilder comment = new StringBuilder();

            // Prepend the object name in the required format
            comment.append("(").append(objectName).append(")- ").append(currentDate).append(": As per the utility report, ");

            Map<String, Set<String>> changes = entry.getValue();

            // Append grouped changes
            if (changes.containsKey("functionsRemoved")) {
                comment.append("the function(s) ").append(String.join(", ", changes.get("functionsRemoved"))).append(" has been removed, ");
            }
            if (changes.containsKey("proceduresAdded")) {
                comment.append("and the procedure(s) ").append(String.join(", ", changes.get("proceduresAdded"))).append(" has been added in the packages, ");
            }
            if (changes.containsKey("columnsModified")) {
                comment.append("the columns ").append(String.join(", ", changes.get("columnsModified"))).append(" have been modified, ");
            }
            if (changes.containsKey("columnsAdded")) {
                comment.append("and the columns (").append(String.join(", ", changes.get("columnsAdded"))).append(") have been added in the tables, ");
            }
            if (changes.containsKey("viewsModified")) {
                comment.append("Also, the views ").append(String.join(", ", changes.get("viewsModified"))).append(" have been modified, ");
            }

            comment.append("in the upgrade instance has no impact on the custom object. Hence, no remediation is required.");
            finalComments.add(comment.toString().replace(", in the upgrade", " in the upgrade"));
        }

        // Log the final comments
        System.out.println("Generated Comments: " + finalComments);

        return finalComments;
    }

            //        List<String> comments = new ArrayList<>(); ====25-03-25====03:25PM=====start
//        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
//        String currentDate = sdf.format(new Date());
//
//        // Map to store consolidated messages per object
//        Map<String, Map<String, Set<String>>> consolidatedComments = new LinkedHashMap<>();
//
//        for (Row row : sheet) {
//            if (row.getRowNum() == 0) continue; // Skip header row
//
//            // Read values
//            String objectName = getCellValue(row, columnIndexes.get("OBJECT_NAME"));
//            String referencedType = getCellValue(row, columnIndexes.get("REFERENCED_TYPE"));
//            String objectSearchName = getCellValue(row, columnIndexes.get("OBJECT_SEARCH_NAME"));
//            String changeType = getCellValue(row, columnIndexes.get("CHANGE_TYPE"));
//            String additionalInfo = getCellValue(row, columnIndexes.get("ADDITIONAL_INFO1_1"));
//
//            // Log values to check if data is being read correctly
//            System.out.println("Reading row " + row.getRowNum() + ": " + objectName + ", " + referencedType + ", " + objectSearchName);
//
//            if (objectSearchName.isEmpty() || changeType.isEmpty()) continue; // Skip empty rows
//
//            // Initialize object if not present
//            consolidatedComments.putIfAbsent(objectName, new LinkedHashMap<>());
//
//            // Identify change category
//            String keyCategory = "";
//            if (referencedType.equalsIgnoreCase("PACKAGE")) {
//                keyCategory = changeType.contains("ADDED") ? "proceduresAdded" : "functionsRemoved";
//            } else if (referencedType.equalsIgnoreCase("TABLE")) {
//                keyCategory = changeType.contains("ADDED") ? "columnsAdded" : "columnsModified";
//            } else if (referencedType.equalsIgnoreCase("VIEW")) {
//                keyCategory = "viewsModified";
//            }
//
//            // Store values in the respective category
//            consolidatedComments.get(objectName).putIfAbsent(keyCategory, new LinkedHashSet<>());
//            consolidatedComments.get(objectName).get(keyCategory).add(objectSearchName + (additionalInfo.isEmpty() ? "" : " (" + additionalInfo + ")"));
//        }
//
//        // **Generating final consolidated comments**
//        List<String> finalComments = new ArrayList<>();
//        for (Map.Entry<String, Map<String, Set<String>>> entry : consolidatedComments.entrySet()) {
//            StringBuilder comment = new StringBuilder();
//            comment.append(currentDate).append(": As per the utility report, ");
//
//            Map<String, Set<String>> changes = entry.getValue();
//
//            // Append grouped changes
//            if (changes.containsKey("functionsRemoved")) {
//                comment.append("the function(s) ").append(String.join(", ", changes.get("functionsRemoved"))).append(" has been removed, ");
//            }
//            if (changes.containsKey("proceduresAdded")) {
//                comment.append("and the procedure(s) ").append(String.join(", ", changes.get("proceduresAdded"))).append(" has been added in the packages, ");
//            }
//            if (changes.containsKey("columnsModified")) {
//                comment.append("the columns ").append(String.join(", ", changes.get("columnsModified"))).append(" have been modified, ");
//            }
//            if (changes.containsKey("columnsAdded")) {
//                comment.append("and the columns (").append(String.join(", ", changes.get("columnsAdded"))).append(") have been added in the tables, ");
//            }
//            if (changes.containsKey("viewsModified")) {
//                comment.append("Also, the views ").append(String.join(", ", changes.get("viewsModified"))).append(" have been modified, ");
//            }
//
//            comment.append("in the upgrade instance has no impact on the custom object. Hence, no remediation is required.");
//            finalComments.add(comment.toString().replace(", in the upgrade", " in the upgrade"));
//        }
//
//        // Log the final comments
//        System.out.println("Generated Comments: " + finalComments);
//
//        return finalComments;
//    } ====25-03-25====03:25PM=====END

    // Read Cell Value Safely
    private static String getCellValue(Row row, Integer cellIndex) {
        if (cellIndex == null) return "";
        Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue()).trim();
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    //  Save Comments to a File
    private static void saveCommentsToFile(List<String> comments, String filePath) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            for (String comment : comments) {
                writer.write(comment);
                writer.newLine();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}

//    private static final Logger logger = LoggerFactory.getLogger(ExcelReaderService.class);
//
//    @Value("${input.file.path}")
//    private String filePath;
//
//    @Value("${output.file.path}")
//    private String outputFilePath;
//
//    public void readDBObjectsSheet() {
//        logger.info("Starting to process DB Objects sheet...");
//        generateComments();
//    }
//
//    public void generateComments() {
//
//
//        try (FileInputStream fis = new FileInputStream(filePath);
//             Workbook workbook = new XSSFWorkbook(fis);
//             BufferedWriter writer = new BufferedWriter(new FileWriter(outputFilePath))) {
//
//            Sheet sheet = workbook.getSheet("DB Objects");
//            if (sheet == null) {
//                logger.error("Error: Sheet 'DB Objects' not found.");
//                writer.write("Error: Sheet 'DB Objects' not found.\n");
//                return;
//            }
//
//            Map<String, List<Map<String, String>>> groupedData = new LinkedHashMap<>();
//
//            for (Row row : sheet) {
//                if (row.getRowNum() == 0) continue; // Skip header row
//                Map<String, String> record = extractRowData(row);
//                groupedData.computeIfAbsent(record.get("date"), k -> new ArrayList<>()).add(record);
//            }
//
//            for (Map.Entry<String, List<Map<String, String>>> entry : groupedData.entrySet()) {
//                writer.write(formatComment(entry.getKey(), entry.getValue()));
//                writer.newLine();
//            }
//
//            logger.info("Excel processing complete. Output saved at: {}", outputFilePath);
//
//        } catch (IOException e) {
//            logger.error("Error processing file: {}", filePath, e);
//        }
//    }
//
//    private Map<String, String> extractRowData(Row row) {
//        Map<String, String> record = new HashMap<>();
//        record.put("date", getCellValue(row, 11));
//        record.put("objectName", getCellValue(row, 3));
//        record.put("referencedType", getCellValue(row, 6));
//        record.put("objectSearchName", getCellValue(row, 7));
//        record.put("changeType", getCellValue(row, 13));
//        record.put("additionalInfo", getCellValue(row, 14));
//        return record;
//    }
//
//    private String formatComment(String date, List<Map<String, String>> records) {
//        Map<String, List<String>> categorizedChanges = new LinkedHashMap<>();
//
//        for (Map<String, String> record : records) {
//            categorizedChanges
//                    .computeIfAbsent(record.get("changeType"), k -> new ArrayList<>())
//                    .add(record.get("objectSearchName"));
//        }
//
//        StringBuilder comment = new StringBuilder();
//        comment.append("Date: ").append(date).append("\nBased on the utility report:\n");
//
//        for (Map.Entry<String, List<String>> entry : categorizedChanges.entrySet()) {
//            comment.append("- ").append(entry.getKey()).append(": ").append(String.join(", ", entry.getValue())).append("\n");
//        }
//
//        comment.append("No impact on the custom object. No remediation required.\n");
//
//        return comment.toString();
//    }
//
//    private String getCellValue(Row row, int cellIndex) {
//        Cell cell = row.getCell(cellIndex);
//        if (cell == null) return "";
//
//        switch (cell.getCellType()) {
//            case STRING:
//                return cell.getStringCellValue().trim();
//            case NUMERIC:
//                if (DateUtil.isCellDateFormatted(cell)) {
//                    return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
//                }
//                return String.valueOf(cell.getNumericCellValue()); // Preserve decimals if present
//            case BOOLEAN:
//                return String.valueOf(cell.getBooleanCellValue());
//            default:
//                return "";
//        }
//    }
//}




//        // Increase the POI memory limit (Set this before loading the Excel file)
//        IOUtils.setByteArrayMaxOverride(300_000_000); // 300MB limit
//
//        // Define File Paths
//        String filePath = "C:\\Users\\Vishal Kumar\\Desktop\\Master_Tracker\\Boral_Master_tracker_v3_detailed.xlsx";
//        String sheetName = "DB Objects";
//        String outputPath = "C:\\Users\\Vishal Kumar\\Desktop\\utility_comments.txt";
//
//        // Check File Exists
//        File file = new File(filePath);
//        if (!file.exists()) {
//            System.err.println(" ERROR: Excel file not found at: " + file.getAbsolutePath());
//            return;
//        }
//
//        try {
//            // Open Excel File
//            FileInputStream fileInputStream = new FileInputStream(file);
//            Workbook workbook = new XSSFWorkbook(fileInputStream);
//            Sheet sheet = workbook.getSheet(sheetName);
//
//            if (sheet == null) {
//                System.err.println(" ERROR: Excel sheet '" + sheetName + "' not found!");
//                workbook.close();
//                return;
//            }
//
//            // Find Column Indexes
//            Map<String, Integer> columnIndexes = getColumnIndexes(sheet);
//            if (columnIndexes.isEmpty()) {
//                System.err.println(" ERROR: Required column headers not found!");
//                workbook.close();
//                return;
//            }
//
//            // Generate Comments
//            List<String> comments = generateComments(sheet, columnIndexes);
//            workbook.close();
//
//            // Save Comments to File
//            saveCommentsToFile(comments, outputPath);
//            System.out.println(" Comments saved to: " + outputPath);
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//
//    // Identify Column Indexes Dynamically
//    private static Map<String, Integer> getColumnIndexes(Sheet sheet) {
//        Map<String, Integer> columnIndexes = new HashMap<>();
//
//        Row headerRow = sheet.getRow(1);
//        if (headerRow == null) {
//            System.err.println("ERROR: Header row is empty!");
//            return columnIndexes;
//        }
//
//        System.out.println("🔍 Headers found in Row :");
//        for (Cell cell : headerRow) {
//            //String header = cell.getStringCellValue().trim();
//            String header = getCellValue(headerRow, cell.getColumnIndex());
//            System.out.println(" - " + header);
//
//
//            switch (header) {
//                case "OBJECT_NAME":
//                case "REFERENCED_TYPE":
//                case "OBJECT_SEARCH_NAME":
//                case "CHANGE_TYPE":
//                case "ADDITIONAL_INFO1_1":
//                    columnIndexes.put(header, cell.getColumnIndex());
//                    break;
//            }
//        }
//
//
//        return columnIndexes;
//    }
//
//    // Generate Comments Based on Utility Report
//    private static List<String> generateComments(Sheet sheet, Map<String, Integer> columnIndexes) {
//        List<String> comments = new ArrayList<>();
//        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
//        String currentDate = sdf.format(new Date());
//
//        // Map to store consolidated messages per object
//        Map<String, Map<String, Set<String>>> consolidatedComments = new LinkedHashMap<>();
//
//        for (Row row : sheet) {
//            if (row.getRowNum() == 0) continue; // Skip header row
//
              // Read values
//            String objectName = getCellValue(row, columnIndexes.get("OBJECT_NAME"));
//            String referencedType = getCellValue(row, columnIndexes.get("REFERENCED_TYPE"));
//            String objectSearchName = getCellValue(row, columnIndexes.get("OBJECT_SEARCH_NAME"));
//            String changeType = getCellValue(row, columnIndexes.get("CHANGE_TYPE"));
//            String additionalInfo = getCellValue(row, columnIndexes.get("ADDITIONAL_INFO1_1"));
//
//            if (objectSearchName.isEmpty() || changeType.isEmpty()) continue; // Skip empty rows
//
//            // Initialize object if not present
//            consolidatedComments.putIfAbsent(objectName, new LinkedHashMap<>());
//
//            // Identify change category
//            String keyCategory = "";
//            if (referencedType.equalsIgnoreCase("PACKAGE")) {
//                keyCategory = changeType.contains("ADDED") ? "proceduresAdded" : "functionsRemoved";
//            } else if (referencedType.equalsIgnoreCase("TABLE")) {
//                keyCategory = changeType.contains("ADDED") ? "columnsAdded" : "columnsModified";
//            } else if (referencedType.equalsIgnoreCase("VIEW")) {
//                keyCategory = "viewsModified";
//            }
//
//            // Store values in the respective category
//            consolidatedComments.get(objectName).putIfAbsent(keyCategory, new LinkedHashSet<>());
//            consolidatedComments.get(objectName).get(keyCategory).add(objectSearchName + (additionalInfo.isEmpty() ? "" : " (" + additionalInfo + ")"));
//        }
//
//        // **Generating final consolidated comments**
//        List<String> finalComments = new ArrayList<>();
//        for (Map.Entry<String, Map<String, Set<String>>> entry : consolidatedComments.entrySet()) {
//            StringBuilder comment = new StringBuilder();
//            comment.append(currentDate).append(": As per the utility report, ");
//
//            Map<String, Set<String>> changes = entry.getValue();
//
//            // Append grouped changes
//            if (changes.containsKey("functionsRemoved")) {
//                comment.append("the function(s) ").append(String.join(", ", changes.get("functionsRemoved"))).append(" has been removed, ");
//            }
//            if (changes.containsKey("proceduresAdded")) {
//                comment.append("and the procedure(s) ").append(String.join(", ", changes.get("proceduresAdded"))).append(" has been added in the packages, ");
//            }
//            if (changes.containsKey("columnsModified")) {
//                comment.append("the columns ").append(String.join(", ", changes.get("columnsModified"))).append(" have been modified, ");
//            }
//            if (changes.containsKey("columnsAdded")) {
//                comment.append("and the columns (").append(String.join(", ", changes.get("columnsAdded"))).append(") have been added in the tables, ");
//            }
//            if (changes.containsKey("viewsModified")) {
//                comment.append("Also, the views ").append(String.join(", ", changes.get("viewsModified"))).append(" have been modified, ");
//            }
//
//            comment.append("in the upgrade instance has no impact on the custom object. Hence, no remediation is required.");
//            finalComments.add(comment.toString().replace(", in the upgrade", " in the upgrade"));
//        }
//
//        return finalComments;
//    }
//
//    // Read Cell Value Safely
//    private static String getCellValue(Row row, Integer cellIndex) {
//        if (cellIndex == null) return "";
//        Cell cell = row.getCell(cellIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
//        if (cell == null) return "";
//
//        switch (cell.getCellType()) {
//            case STRING:
//                return cell.getStringCellValue().trim();
//            case NUMERIC:
//                return String.valueOf((int) cell.getNumericCellValue()).trim();
//            case BOOLEAN:
//                return String.valueOf(cell.getBooleanCellValue());
//            default:
//                return "";
//        }
//    }
//
//    //  Save Comments to a File
//    private static void saveCommentsToFile(List<String> comments, String filePath) {
//        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
//            for (String comment : comments) {
//                writer.write(comment);
//                writer.newLine();
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    public void readDBObjectsSheet() {
 // }

