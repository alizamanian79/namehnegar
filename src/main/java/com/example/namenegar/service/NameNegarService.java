package com.example.namenegar.service;

import org.springframework.http.ResponseEntity;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;
import java.util.Map;

public interface NameNegarService {
    List<Map<String, Object>> readExcelAsMap(MultipartFile file);
    String generateList(MultipartFile file);
    List<String> generateWordFilesFromExcelAndTemplate(MultipartFile excelFile, MultipartFile wordTemplate);
    ResponseEntity<byte[]> downloadAllFiles();
    List<String> filesList();

}
