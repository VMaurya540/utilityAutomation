# Utility Comments Automation

A Java 21 and Spring Boot app that builds Oracle object comments from a master Excel tracker and writes them into a developer tracker. It provides a local web form at `http://localhost:8080` and a JSON API.

The Maven project is in the [`utilityAutomation`](utilityAutomation/) folder. SharePoint integration is currently a placeholder; it does not download or upload files.

## Requirements

- JDK 21
- Maven 3.9+ (or the Maven bundled with IntelliJ IDEA)
- Two `.xlsx` files: a master tracker and the tracker to update. These files are not included in this repository.

## Run locally

Open the `utilityAutomation` folder as a Maven project in IntelliJ IDEA. Set **Project SDK** to JDK 21 and wait for Maven import to finish. Run the Maven goal `spring-boot:run`, or run this in a terminal from the repository root:

```powershell
cd utilityAutomation
mvn spring-boot:run
```

Open [http://localhost:8080](http://localhost:8080). The home page loads without Excel files. To process data, replace the example paths in **Master Tracker Path** and **Tracker Path** with real files on your computer. The sample Windows paths shown in the form are placeholders from the original project.

Check the server with `GET http://localhost:8080/api/health`; a running app returns `OK`.

## Local automation

The app reads the master workbook, groups changes by object name and type, generates comments, and updates the tracker's `ORACLE COMMENTS` column. It filters tracker rows by `IDENTIFIED BY` and allowed object types. Sheet names can be entered explicitly or set to `auto` to find a sheet by its headers.

The master sheet needs `OBJECT_NAME`, `OBJECT TYPE`, `OBJECT_SEARCH_NAME`, `REFERENCED_TYPE`, `CHANGE_TYPE`, and `ADDITIONAL_INFO1_1`. The tracker sheet needs `OBJECT_NAME`, `OBJECT TYPE`, `IDENTIFIED BY`, and `ORACLE COMMENTS`.

By default, **Output Tracker Path** is blank, so a successful run overwrites the source tracker. Set an output path to create a separate updated workbook. Make a backup before running on important data. The app writes a run summary log and, when some objects have no generated comment, a missed objects log. **Output Log Directory** controls where those logs go; otherwise they go beside the output tracker.

Default allowed object types are `Package`, `Package Body`, `Procedure`, and `View`; the default `Identified By` filter is `Utility`. The form allows these and the date format to be changed for a run. **Fallback Object Name Only** allows a match by name when a name and type match is not found.

## API

`POST /api/local-run` accepts JSON with these fields:

```json
{
  "masterPath": "C:\\path\\to\\master.xlsx",
  "trackerPath": "C:\\path\\to\\tracker.xlsx",
  "outputTrackerPath": "C:\\path\\to\\updated-tracker.xlsx",
  "masterSheetName": "auto",
  "trackerSheetName": "auto",
  "identifiedBy": "Utility",
  "allowedObjectTypes": ["Package", "Package Body", "Procedure", "View"],
  "dateFormat": "dd-MMM-yyyy",
  "fallbackObjectNameOnly": true,
  "outputDir": "C:\\path\\to\\logs"
}
```

The response includes `success`, `message`, row counts, the tracker output path, and any generated log paths.

`POST /api/sharepoint-run` currently returns a message describing the planned Microsoft Graph integration and does not process a workbook.

## Build and test

From the `utilityAutomation` folder:

```powershell
mvn test
mvn package
```

Keep real trackers, credentials, and generated logs out of Git. No workbook or SharePoint credentials are required to start the web app.
