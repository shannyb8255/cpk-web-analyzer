package com.cpk.cpk_web_analyzer;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

@Controller
public class ExcelController {

    @GetMapping("/")
    public String showUploadForm() {
        return "upload";
    }

    @PostMapping("/analyze")
    public String handleFileUpload(@RequestParam("file") MultipartFile file, Model model) {
        if (file.isEmpty()) {
            model.addAttribute("message", "⚠️ Please select a file to upload.");
            return "upload";
        }

        try {
            // Save uploaded file to temp directory
            File tempFile = File.createTempFile("uploaded-", ".xlsx");
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                fos.write(file.getBytes());
            }

            // Run your processing logic
            ExcelCapabilityTemplate processor = new ExcelCapabilityTemplate();
            processor.processFile(tempFile);

            model.addAttribute("message", "✅ Analysis complete! File processed successfully.");
        } catch (IOException e) {
            model.addAttribute("message", "❌ Error processing file: " + e.getMessage());
        }

        return "upload";
    }
}
