package com.example.namenegar.NamehNegar.controller;

import com.example.namenegar.service.NameNegarService;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.net.MalformedURLException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;


@RestController
@RequiredArgsConstructor
@RequestMapping("/api/namehnegar/v1")
public class NameNegarController {


    private final NameNegarService nameNegarService;

    // Objects in excel files
    @PostMapping("/excel-properties")
    public ResponseEntity<List<Map<String, Object>>> uploadExcel(@RequestParam("excelFile") MultipartFile file) {
        try {
            List<Map<String, Object>> result = nameNegarService.readExcelAsMap(file);
            return ResponseEntity.ok(result);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }

    // Generate object list in table
    @PostMapping("/personList")
    public ResponseEntity<Map<String, Object>> convertExcelToPdf(@RequestParam("file") MultipartFile file) {
        try {
            String filePath = nameNegarService.generateList(file);
            return ResponseEntity.ok(Map.of(
                    "message", "فایل PDF با موفقیت ساخته شد.",
                    "path", filePath
            ));
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(
                    Map.of("error", "خطا در تبدیل اکسل به PDF: " + e.getMessage())
            );
        }
    }

    // تولید فایل‌های Word از اکسل و قالب ورد
    @PostMapping("/word/generate")
    public ResponseEntity<?> generateWordFiles(
            @RequestParam("excelFile") MultipartFile excelFile,
            @RequestParam("wordTemplate") MultipartFile wordTemplate) {
        try {
            List<String> generatedFiles = nameNegarService.generateWordFilesFromExcelAndTemplate(excelFile, wordTemplate);
            return ResponseEntity.ok(generatedFiles);
        } catch (Exception e) {
            e.printStackTrace(); // چاپ دقیق خطا در لاگ سرور
            return ResponseEntity.badRequest().body("خطا: " + e.getMessage());
        }
    }


    @GetMapping("/word/download-word")
    public ResponseEntity<Resource> downloadWordFile(@RequestParam String filename) {
        try {
            Path filePath = Paths.get("uploads/word/").resolve(filename).normalize();
            Resource resource = new UrlResource(filePath.toUri());

            if (resource.exists()) {
                HttpHeaders headers = new HttpHeaders();
                headers.add(HttpHeaders.CONTENT_DISPOSITION, "inline; filename=\"" + resource.getFilename() + "\"");

                return ResponseEntity.ok()
                        .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                        .headers(headers)
                        .body(resource);
            } else {
                return ResponseEntity.status(HttpStatus.NOT_FOUND).body(null);
            }
        } catch (MalformedURLException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(null);
        }
    }


    @GetMapping("/word/files/list")
    public List<String> listFilesWithPasswordSuffix() {
        return nameNegarService.filesList();
    }


    public static List<String> list() {
        String folderPath = "uploads/word"; // مسیر فولدر

        List<String> result = new ArrayList<>();
        File folder = new File(folderPath);

        if (folder.exists() && folder.isDirectory()) {
            File[] files = folder.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isFile()) {
                        String filename = file.getName();
                        result.add("http://localhost:8080/api/v1/download-word?filename=" +filename);
                    }
                }
            }
        }

        return result;
    }
    @GetMapping("/word/files/list/download")
    public ResponseEntity<byte[]> downloadAllFiles() {
        return nameNegarService.downloadAllFiles();
    }




}
