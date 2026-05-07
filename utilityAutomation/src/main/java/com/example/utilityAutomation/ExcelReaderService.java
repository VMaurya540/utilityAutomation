package com.example.utilityAutomation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

@Service
public class ExcelReaderService {

    private static final String DEFAULT_MASTER_TRACKER_PATH =
            "C:\\Projects\\boral\\Boral_Retrofitment\\Utility\\Final_Utility_28-02-2025\\Boral_Master_tracker_v3_detailed.xlsx";
    private static final String DEFAULT_DEVELOPER_TRACKER_PATH =
            "C:\\Projects\\boral\\Boral_Retrofitment\\Utility\\Final_Utility_28-02-2025\\Boral_RICE_CEMLI_Tracker_V3.xlsx";

    private static final String MASTER_SHEET_DB_OBJECTS = "DB Objects";
    private static final String TRACKER_SHEET_NAME = "Tracker";

    private static final Set<String> UTILITY_OBJECT_TYPES = new HashSet<String>(Arrays.asList(
            "PACKAGE", "PACKAGE BODY", "PROCEDURE", "VIEW"
    ));

    private static final String IDENTIFIED_BY_UTILITY = "UTILITY";
    private static final DataFormatter DATA_FORMATTER = new DataFormatter();

    public void readDBObjectsSheet() {
        IOUtils.setByteArrayMaxOverride(350_000_000);

        String masterTrackerPath = System.getProperty("utility.master.path", DEFAULT_MASTER_TRACKER_PATH);
        String developerTrackerPath = System.getProperty("utility.tracker.path", DEFAULT_DEVELOPER_TRACKER_PATH);

        File masterFile = new File(masterTrackerPath);
        File trackerFile = new File(developerTrackerPath);

        if (!masterFile.exists()) {
            System.err.println("ERROR: Master tracker not found at: " + masterFile.getAbsolutePath());
            return;
        }

        if (!trackerFile.exists()) {
            System.err.println("ERROR: Developer tracker not found at: " + trackerFile.getAbsolutePath());
            return;
        }

        try {
            Map<ObjectKey, String> commentsByObjectAndType = buildUtilityComments(masterFile);
            if (commentsByObjectAndType.isEmpty()) {
                System.out.println("No utility comments generated from master tracker.");
                return;
            }

            TrackerUpdateStats stats = updateTrackerOracleComments(trackerFile, commentsByObjectAndType);

            System.out.println("Utility comments update completed.");
            System.out.println("Rows updated: " + stats.updatedRows);
            System.out.println("Rows skipped (no comment found): " + stats.skippedNoCommentRows);
            System.out.println("Rows skipped (filters): " + stats.skippedFilteredRows);
            if (!isBlank(stats.missedLogFilePath)) {
                System.out.println("Missed objects log: " + stats.missedLogFilePath);
            }

        } catch (Exception e) {
            System.err.println("ERROR while updating utility comments: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private Map<ObjectKey, String> buildUtilityComments(File masterFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(masterFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(MASTER_SHEET_DB_OBJECTS);
            if (sheet == null) {
                throw new IllegalStateException("Sheet not found in master tracker: " + MASTER_SHEET_DB_OBJECTS);
            }

            Set<String> requiredHeaders = new HashSet<String>(Arrays.asList(
                    "OBJECT_NAME",
                    "OBJECT TYPE",
                    "OBJECT_SEARCH_NAME",
                    "REFERENCED_TYPE",
                    "CHANGE_TYPE",
                    "ADDITIONAL_INFO1_1"
            ));

            HeaderInfo headerInfo = detectHeader(sheet, requiredHeaders);
            if (headerInfo == null) {
                throw new IllegalStateException("Required headers not found in master tracker sheet: " + MASTER_SHEET_DB_OBJECTS);
            }

            Map<ObjectKey, Map<String, LinkedHashSet<String>>> groupedChanges = new LinkedHashMap<ObjectKey, Map<String, LinkedHashSet<String>>>();

            for (int r = headerInfo.rowIndex + 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }

                String objectName = getCellValue(row, headerInfo.columns.get("OBJECT_NAME"));
                String objectType = normalizeObjectType(getCellValue(row, headerInfo.columns.get("OBJECT TYPE")));
                String referencedType = normalizeObjectType(getCellValue(row, headerInfo.columns.get("REFERENCED_TYPE")));
                String objectSearchName = getCellValue(row, headerInfo.columns.get("OBJECT_SEARCH_NAME"));
                String changeType = normalizeToken(getCellValue(row, headerInfo.columns.get("CHANGE_TYPE")));
                String additionalInfo = getCellValue(row, headerInfo.columns.get("ADDITIONAL_INFO1_1"));

                if (isBlank(objectName) || isBlank(objectType) || isBlank(objectSearchName)) {
                    continue;
                }

                if (!UTILITY_OBJECT_TYPES.contains(objectType)) {
                    continue;
                }

                String category = getCategory(referencedType, changeType);
                if (isBlank(category)) {
                    continue;
                }

                ObjectKey key = new ObjectKey(normalizeToken(objectName), objectType);
                if (!groupedChanges.containsKey(key)) {
                    groupedChanges.put(key, new LinkedHashMap<String, LinkedHashSet<String>>());
                }
                if (!groupedChanges.get(key).containsKey(category)) {
                    groupedChanges.get(key).put(category, new LinkedHashSet<String>());
                }

                String value = isBlank(additionalInfo)
                        ? objectSearchName
                        : objectSearchName + " (" + additionalInfo + ")";

                groupedChanges.get(key).get(category).add(value);
            }

            SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy", Locale.ENGLISH);
            String currentDate = sdf.format(new Date());

            Map<ObjectKey, String> finalComments = new HashMap<ObjectKey, String>();
            for (Map.Entry<ObjectKey, Map<String, LinkedHashSet<String>>> entry : groupedChanges.entrySet()) {
                finalComments.put(entry.getKey(), buildComment(currentDate, entry.getValue()));
            }

            return finalComments;
        }
    }

    private TrackerUpdateStats updateTrackerOracleComments(File trackerFile,
                                                           Map<ObjectKey, String> commentsByObjectAndType) throws IOException {

        try (FileInputStream fis = new FileInputStream(trackerFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet trackerSheet = workbook.getSheet(TRACKER_SHEET_NAME);
            if (trackerSheet == null) {
                throw new IllegalStateException("Sheet not found in tracker: " + TRACKER_SHEET_NAME);
            }

            Set<String> requiredHeaders = new HashSet<String>(Arrays.asList(
                    "OBJECT_NAME",
                    "OBJECT TYPE",
                    "IDENTIFIED BY",
                    "ORACLE COMMENTS"
            ));

            HeaderInfo headerInfo = detectHeader(trackerSheet, requiredHeaders);
            if (headerInfo == null) {
                throw new IllegalStateException("Required headers not found in tracker sheet: " + TRACKER_SHEET_NAME);
            }

            int objectNameCol = headerInfo.columns.get("OBJECT_NAME").intValue();
            int objectTypeCol = headerInfo.columns.get("OBJECT TYPE").intValue();
            int identifiedByCol = headerInfo.columns.get("IDENTIFIED BY").intValue();
            int oracleCommentsCol = headerInfo.columns.get("ORACLE COMMENTS").intValue();

            TrackerUpdateStats stats = new TrackerUpdateStats();
            List<MissedObjectEntry> missedObjects = new ArrayList<MissedObjectEntry>();

            for (int r = headerInfo.rowIndex + 1; r <= trackerSheet.getLastRowNum(); r++) {
                Row row = trackerSheet.getRow(r);
                if (row == null) {
                    continue;
                }

                String identifiedBy = normalizeToken(getCellValue(row, Integer.valueOf(identifiedByCol)));
                String objectType = normalizeObjectType(getCellValue(row, Integer.valueOf(objectTypeCol)));
                String objectName = normalizeToken(getCellValue(row, Integer.valueOf(objectNameCol)));

                if (!IDENTIFIED_BY_UTILITY.equals(identifiedBy)
                        || !UTILITY_OBJECT_TYPES.contains(objectType)
                        || isBlank(objectName)) {
                    stats.skippedFilteredRows++;
                    continue;
                }

                ObjectKey key = new ObjectKey(objectName, objectType);
                String comment = commentsByObjectAndType.get(key);

                if (isBlank(comment)) {
                    stats.skippedNoCommentRows++;
                    missedObjects.add(new MissedObjectEntry(r + 1, objectName, objectType, "No generated utility comment found from DB Objects"));
                    continue;
                }

                Cell commentCell = row.getCell(oracleCommentsCol, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                commentCell.setCellValue(comment);
                stats.updatedRows++;
            }

            try (FileOutputStream fos = new FileOutputStream(trackerFile)) {
                workbook.write(fos);
            }
            stats.missedLogFilePath = writeMissedObjectsLog(trackerFile, missedObjects);

            return stats;
        }
    }

    private String writeMissedObjectsLog(File trackerFile, List<MissedObjectEntry> missedObjects) throws IOException {
        if (missedObjects == null || missedObjects.isEmpty()) {
            return "";
        }

        String ts = new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.ENGLISH).format(new Date());
        File output = new File(trackerFile.getParentFile(), "utility_missed_objects_" + ts + ".log");

        try (BufferedWriter writer = new BufferedWriter(new FileWriter(output))) {
            writer.write("Missed Utility Objects");
            writer.newLine();
            writer.write("Generated At: " + new SimpleDateFormat("dd-MMM-yyyy HH:mm:ss", Locale.ENGLISH).format(new Date()));
            writer.newLine();
            writer.write("Tracker: " + trackerFile.getAbsolutePath());
            writer.newLine();
            writer.newLine();

            for (MissedObjectEntry entry : missedObjects) {
                writer.write("Row=" + entry.rowNumber
                        + " | Object_name=" + entry.objectName
                        + " | Object Type=" + entry.objectType
                        + " | Reason=" + entry.reason);
                writer.newLine();
            }
        }

        return output.getAbsolutePath();
    }

    private HeaderInfo detectHeader(Sheet sheet, Set<String> requiredHeaders) {
        for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 25); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }

            Map<String, Integer> found = new HashMap<String, Integer>();
            for (Cell cell : row) {
                String rawHeader = DATA_FORMATTER.formatCellValue(cell).trim();
                if (isBlank(rawHeader)) {
                    continue;
                }

                String normalized = normalizeHeader(rawHeader);

                if (normalized.equals("OBJECTNAME") || normalized.equals("OBJECT_NAME")) {
                    found.put("OBJECT_NAME", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("OBJECTTYPE") || normalized.equals("OBJECT_TYPE")) {
                    found.put("OBJECT TYPE", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("IDENTIFIEDBY") || normalized.equals("IDENTIFIED_BY")) {
                    found.put("IDENTIFIED BY", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("ORACLECOMMENTS") || normalized.equals("ORACLE_COMMENTS")) {
                    found.put("ORACLE COMMENTS", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("OBJECTSEARCHNAME") || normalized.equals("OBJECT_SEARCH_NAME")) {
                    found.put("OBJECT_SEARCH_NAME", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("REFERENCEDTYPE") || normalized.equals("REFERENCED_TYPE")) {
                    found.put("REFERENCED_TYPE", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("CHANGETYPE") || normalized.equals("CHANGE_TYPE")) {
                    found.put("CHANGE_TYPE", Integer.valueOf(cell.getColumnIndex()));
                } else if (normalized.equals("ADDITIONALINFO11") || normalized.equals("ADDITIONAL_INFO1_1")) {
                    found.put("ADDITIONAL_INFO1_1", Integer.valueOf(cell.getColumnIndex()));
                }
            }

            if (found.keySet().containsAll(requiredHeaders)) {
                return new HeaderInfo(r, found);
            }
        }

        return null;
    }

    private String buildComment(String currentDate, Map<String, LinkedHashSet<String>> changes) {
        List<String> fragments = new ArrayList<String>();

        if (changes.containsKey("functionsRemoved")) {
            fragments.add("the function(s) " + joinValues(changes.get("functionsRemoved")) + " have been removed");
        }

        if (changes.containsKey("proceduresAdded")) {
            fragments.add("the procedure(s) " + joinValues(changes.get("proceduresAdded")) + " have been added");
        }

        if (changes.containsKey("proceduresModified")) {
            fragments.add("the procedure(s) " + joinValues(changes.get("proceduresModified")) + " have been modified");
        }

        if (changes.containsKey("columnsModified")) {
            fragments.add("the columns " + joinValues(changes.get("columnsModified")) + " have been modified");
        }

        if (changes.containsKey("columnsAdded")) {
            fragments.add("the columns " + joinValues(changes.get("columnsAdded")) + " have been added");
        }

        if (changes.containsKey("viewsModified")) {
            fragments.add("the views " + joinValues(changes.get("viewsModified")) + " have been modified");
        }

        if (fragments.isEmpty()) {
            return currentDate + ": As per utility report, standard objects were reviewed in the upgrade instance and have no impact on the custom object. Hence, no remediation is required.";
        }

        return currentDate + ": As per utility report, "
                + joinValues(fragments)
                + " in the upgrade instance has no impact on the custom object. Hence, no remediation is required.";
    }

    private String getCategory(String referencedType, String changeType) {
        if (isBlank(referencedType) || isBlank(changeType)) {
            return "";
        }

        if (referencedType.equals("PACKAGE") || referencedType.equals("PACKAGE BODY")
                || referencedType.equals("PROCEDURE") || referencedType.equals("FUNCTION")) {

            if (changeType.contains("FUNC_REMOVED") || changeType.contains("FUNCTION_REMOVED")) {
                return "functionsRemoved";
            }

            if (changeType.contains("PROC_ADDED") || changeType.contains("PROCEDURE_ADDED") || changeType.contains("ADDED")) {
                return "proceduresAdded";
            }

            if (changeType.contains("PROC_MODIFIED") || changeType.contains("PROCEDURE_MODIFIED") || changeType.contains("MODIFIED")) {
                return "proceduresModified";
            }
        }

        if (referencedType.equals("TABLE")) {
            if (changeType.contains("ADDED")) {
                return "columnsAdded";
            }
            if (changeType.contains("MODIFIED") || changeType.contains("CHANGED")) {
                return "columnsModified";
            }
        }

        if (referencedType.equals("VIEW")) {
            if (changeType.contains("MODIFIED") || changeType.contains("CHANGED") || changeType.contains("VIEW")) {
                return "viewsModified";
            }
        }

        return "";
    }

    private String getCellValue(Row row, Integer cellIndex) {
        if (cellIndex == null) {
            return "";
        }

        Cell cell = row.getCell(cellIndex.intValue(), MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) {
            return "";
        }

        if (cell.getCellType() == CellType.FORMULA) {
            return DATA_FORMATTER.formatCellValue(cell).trim();
        }

        return DATA_FORMATTER.formatCellValue(cell).trim();
    }

    private String normalizeHeader(String value) {
        return value == null
                ? ""
                : value.replace("_", "")
                .replace(" ", "")
                .toUpperCase(Locale.ENGLISH)
                .trim();
    }

    private String normalizeToken(String value) {
        return value == null ? "" : value.trim().toUpperCase(Locale.ENGLISH);
    }

    private String normalizeObjectType(String objectType) {
        String normalized = normalizeToken(objectType).replaceAll("\\s+", " ");

        if (normalized.equals("PACKAGEBODY")) {
            return "PACKAGE BODY";
        }

        return normalized;
    }

    private String joinValues(Iterable<String> values) {
        StringBuilder sb = new StringBuilder();
        boolean first = true;
        for (String v : values) {
            if (!first) {
                sb.append(", ");
            }
            sb.append(v);
            first = false;
        }
        return sb.toString();
    }

    private boolean isBlank(String value) {
        return value == null || value.trim().isEmpty();
    }

    private static class HeaderInfo {
        private final int rowIndex;
        private final Map<String, Integer> columns;

        private HeaderInfo(int rowIndex, Map<String, Integer> columns) {
            this.rowIndex = rowIndex;
            this.columns = columns;
        }
    }

    private static class TrackerUpdateStats {
        private int updatedRows;
        private int skippedNoCommentRows;
        private int skippedFilteredRows;
        private String missedLogFilePath;
    }

    private static class MissedObjectEntry {
        private final int rowNumber;
        private final String objectName;
        private final String objectType;
        private final String reason;

        private MissedObjectEntry(int rowNumber, String objectName, String objectType, String reason) {
            this.rowNumber = rowNumber;
            this.objectName = objectName;
            this.objectType = objectType;
            this.reason = reason;
        }
    }

    private static class ObjectKey {
        private final String objectName;
        private final String objectType;

        private ObjectKey(String objectName, String objectType) {
            this.objectName = objectName;
            this.objectType = objectType;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) {
                return true;
            }
            if (o == null || getClass() != o.getClass()) {
                return false;
            }
            ObjectKey objectKey = (ObjectKey) o;
            return Objects.equals(objectName, objectKey.objectName)
                    && Objects.equals(objectType, objectKey.objectType);
        }

        @Override
        public int hashCode() {
            return Objects.hash(objectName, objectType);
        }
    }
}
