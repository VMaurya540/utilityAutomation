package com.example.utilityAutomation;

import com.example.utilityAutomation.dto.LocalRunRequest;
import com.example.utilityAutomation.dto.RunResponse;
import com.example.utilityAutomation.dto.SharePointRunRequest;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api")
public class UtilityAutomationController {

    private final ExcelReaderService excelReaderService;
    private final SharePointAutomationService sharePointAutomationService;

    public UtilityAutomationController(ExcelReaderService excelReaderService,
                                       SharePointAutomationService sharePointAutomationService) {
        this.excelReaderService = excelReaderService;
        this.sharePointAutomationService = sharePointAutomationService;
    }

    @GetMapping("/health")
    public String health() {
        return "OK";
    }

    @PostMapping(value = "/local-run", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
    public RunResponse runLocal(@RequestBody LocalRunRequest request) {
        return excelReaderService.runLocal(request);
    }

    @PostMapping(value = "/sharepoint-run", consumes = MediaType.APPLICATION_JSON_VALUE, produces = MediaType.APPLICATION_JSON_VALUE)
    public RunResponse runSharePoint(@RequestBody SharePointRunRequest request) {
        return sharePointAutomationService.runSkeleton(request);
    }
}
