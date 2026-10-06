package com.example.utilityAutomation.dto;

import java.util.List;

public class LocalRunRequest {
    public String masterPath;
    public String trackerPath;
    public String outputTrackerPath;
    public String masterSheetName;
    public String trackerSheetName;
    public String identifiedBy;
    public List<String> allowedObjectTypes;
    public String dateFormat;
    public boolean fallbackObjectNameOnly = true;
    public String outputDir;
}
