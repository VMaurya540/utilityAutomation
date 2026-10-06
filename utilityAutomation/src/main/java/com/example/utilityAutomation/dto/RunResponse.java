package com.example.utilityAutomation.dto;

public class RunResponse {
    public boolean success;
    public String message;
    public int updatedRows;
    public int skippedNoCommentRows;
    public int skippedFilteredRows;
    public String trackerOutputPath;
    public String missedLogFilePath;
    public String runSummaryFilePath;

    public static RunResponse failure(String message) {
        RunResponse r = new RunResponse();
        r.success = false;
        r.message = message;
        return r;
    }

    public static RunResponse success(String message) {
        RunResponse r = new RunResponse();
        r.success = true;
        r.message = message;
        return r;
    }
}
