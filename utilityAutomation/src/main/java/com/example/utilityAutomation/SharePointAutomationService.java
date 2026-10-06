package com.example.utilityAutomation;

import com.example.utilityAutomation.dto.RunResponse;
import com.example.utilityAutomation.dto.SharePointRunRequest;
import org.springframework.stereotype.Service;

@Service
public class SharePointAutomationService {

    public RunResponse runSkeleton(SharePointRunRequest request) {
        StringBuilder message = new StringBuilder();
        message.append("SharePoint skeleton is ready. Next implementation steps: ");
        message.append("1) Get Microsoft Graph access token, ");
        message.append("2) Download tracker from drive item/content, ");
        message.append("3) Run local automation on temp file, ");
        message.append("4) Upload updated file back to SharePoint item/content.");
        if (request != null && request.siteUrl != null && !request.siteUrl.trim().isEmpty()) {
            message.append(" Received siteUrl=").append(request.siteUrl.trim()).append(".");
        }
        return RunResponse.failure(message.toString());
    }
}
