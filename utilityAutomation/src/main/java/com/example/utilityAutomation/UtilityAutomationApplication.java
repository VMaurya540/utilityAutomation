package com.example.utilityAutomation;

import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class UtilityAutomationApplication implements CommandLineRunner {

	@Autowired
	private ExcelReaderService excelReaderService;

	public static void main(String[] args) {
		SpringApplication.run(UtilityAutomationApplication.class, args);
	}

	@Override
	public void run(String... args) {
		System.out.println("Starting Excel Reading Process...");
		excelReaderService.readDBObjectsSheet();
		System.out.println("Utility Processing Complete!");
	}

}
