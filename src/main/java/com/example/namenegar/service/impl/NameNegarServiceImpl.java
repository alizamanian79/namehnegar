package com.example.namenegar.service.impl;

import com.example.namenegar.service.NameNegarService;
import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.multipart.MultipartFile;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.apache.poi.xwpf.usermodel.*;

import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.io.*;
import java.nio.file.*;
import java.util.regex.*;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import java.io.InputStream;
import java.util.*;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
@RequiredArgsConstructor
public class NameNegarServiceImpl implements NameNegarService {

    @Value("${app.namenegar.server}")
    private String serverIp;

    @Value("${app.namenegar.word.path}")
    private String wordDir;

    @Value("${app.namenegar.pdf.path}")
    private String pdfDir;

    @Override
    public List<Map<String, Object>> readExcelAsMap(MultipartFile file) {
        List<Map<String, Object>> dataList = new ArrayList<>();

        try (InputStream is = file.getInputStream(); Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            List<String> headers = new ArrayList<>();

            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    headers.add(cell.getStringCellValue());
                }
            }

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Map<String, Object> rowData = new LinkedHashMap<>();

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    rowData.put(headers.get(i), getCellValue(cell));
                }

                dataList.add(rowData);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return dataList;
    }

    private static Object getCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell) ? cell.getDateCellValue() : cell.getNumericCellValue();
            case FORMULA -> cell.getCellFormula();
            case BLANK -> "";
            default -> "";
        };
    }

    @Override
    public String generateList(MultipartFile file) {
        List<Map<String, Object>> data = readExcelAsMap(file);

        if (data.isEmpty()) {
            throw new RuntimeException("فایل اکسل خالی است یا اطلاعات نامعتبر دارد.");
        }

        String uploadDir = pdfDir + "/";
        Path uploadPath = Path.of(uploadDir);
        try {
            if (!Files.exists(uploadPath)) {
                Files.createDirectories(uploadPath);
            }

            String filename = "people_" + System.currentTimeMillis() + ".pdf";
            String filepath = uploadDir + filename;

            Document document = new Document();
            PdfWriter.getInstance(document, new FileOutputStream(filepath));
            document.open();

            Font font = FontFactory.getFont(FontFactory.HELVETICA, 12, BaseColor.BLACK);

            // گرفتن هدرها
            List<String> headers = new ArrayList<>(data.get(0).keySet());

            PdfPTable table = new PdfPTable(headers.size());
            table.setWidthPercentage(100);

            for (String header : headers) {
                PdfPCell cell = new PdfPCell(new Phrase(header, font));
                cell.setBackgroundColor(BaseColor.LIGHT_GRAY);
                table.addCell(cell);
            }

            for (Map<String, Object> row : data) {
                for (String header : headers) {
                    Object val = row.get(header);
                    table.addCell(new Phrase(val != null ? val.toString() : "", font));
                }
            }

            document.add(table);
            document.close();

            return filepath;

        } catch (Exception e) {
            throw new RuntimeException("خطا در ساخت فایل PDF: " + e.getMessage(), e);
        }
    }



    @Override
    public List<String> generateWordFilesFromExcelAndTemplate(MultipartFile excelFile, MultipartFile wordTemplate) {
        List<Map<String, Object>> dataList = readExcelAsMap(excelFile);
        List<String> generatedFiles = new ArrayList<>();

        try {
            Path outputDir = Path.of(wordDir+"/");
            if (!Files.exists(outputDir)) Files.createDirectories(outputDir);

            for (Map<String, Object> data : dataList) {
                try (InputStream templateInputStream = wordTemplate.getInputStream();
                     XWPFDocument document = new XWPFDocument(templateInputStream)) {

                    // تغییر در پاراگراف‌ها
                    for (XWPFParagraph paragraph : document.getParagraphs()) {
                        replacePlaceholdersInParagraph(paragraph, data);
                    }

                    // تغییر در جدول‌ها
                    for (XWPFTable table : document.getTables()) {
                        for (XWPFTableRow row : table.getRows()) {
                            for (XWPFTableCell cell : row.getTableCells()) {
                                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                    replacePlaceholdersInParagraph(paragraph, data);
                                }
                            }
                        }
                    }

                    // گرفتن نام فایل از دو فیلد اول (مثلاً firstName و lastName)
                    String firstKey = data.keySet().stream().toList().get(0);
                    String secondKey = data.keySet().stream().toList().get(1);

                    String firstValue = data.getOrDefault(firstKey, "").toString().replaceAll("\\s+", "_");
                    String secondValue = data.getOrDefault(secondKey, "").toString().replaceAll("\\s+", "_");

                    String filename = firstValue + "_" + secondValue + ".docx";
                    Path filePath = outputDir.resolve(filename);

                    // اگر فایل قبلاً وجود دارد، حذفش کن
                    if (Files.exists(filePath)) {
                        Files.delete(filePath);
                    }

                    try (OutputStream out = Files.newOutputStream(filePath)) {
                        document.write(out);
                    }

                    generatedFiles.add(serverIp + "/api/namehnegar/v1/word/download-word?filename="+filename);

                } catch (Exception e) {
                    e.printStackTrace();
                }
            }

        } catch (Exception e) {
            throw new RuntimeException("خطا در تولید فایل‌های Word: " + e.getMessage(), e);
        }

        return generatedFiles;
    }


    @Override
    public List<String> filesList() {
        String folderPath = wordDir; // مسیر فولدر

        List<String> result = new ArrayList<>();
        File folder = new File(folderPath);

        if (folder.exists() && folder.isDirectory()) {
            File[] files = folder.listFiles();
            if (files != null) {
                for (File file : files) {
                    if (file.isFile()) {
                        String filename = file.getName();
                        result.add(serverIp + "/api/namehnegar/v1/word/download-word?filename=" +filename);
                    }
                }
            }
        }

        return result;
    }

    private void replacePlaceholdersInParagraph(XWPFParagraph paragraph, Map<String, Object> data) {
        StringBuilder fullText = new StringBuilder();

        for (XWPFRun run : paragraph.getRuns()) {
            fullText.append(run.getText(0));
        }

        // جایگزینی متغیرها
        String replacedText = fullText.toString();
        Matcher matcher = Pattern.compile("\\{\\{(.*?)}}").matcher(replacedText);
        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1).trim();
            Object value = data.getOrDefault(key, "");
            matcher.appendReplacement(sb, Matcher.quoteReplacement(value.toString()));
        }
        matcher.appendTail(sb);

        // حذف همه‌ی runهای قبلی
        int runCount = paragraph.getRuns().size();
        for (int i = runCount - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }

        // ایجاد یک run جدید با متن جایگزین‌شده
        XWPFRun newRun = paragraph.createRun();
        newRun.setText(sb.toString());
    }



    @Override
    public ResponseEntity<byte[]> downloadAllFiles() {
        List<String> fileUrls = filesList();
        try {
            RestTemplate restTemplate = new RestTemplate();
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ZipOutputStream zos = new ZipOutputStream(baos);

            for (String url : fileUrls) {
                // درخواست به URL بده و فایل رو بگیر
                byte[] fileBytes = restTemplate.getForObject(url, byte[].class);

                // استخراج نام فایل از query string
                String filename = extractFilenameFromUrl(url);

                // اضافه به zip
                zos.putNextEntry(new ZipEntry(filename));
                zos.write(fileBytes);
                zos.closeEntry();
            }

            zos.close();

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=namenegar-words.zip")
                    .contentType(MediaType.APPLICATION_OCTET_STREAM)
                    .body(baos.toByteArray());

        } catch (Exception e) {
            return ResponseEntity.internalServerError().body(("خطا: " + e.getMessage()).getBytes());
        }
    }


    private String extractFilenameFromUrl(String url) {
        try {
            String[] parts = url.split("filename=");
            if (parts.length > 1) {
                return URLDecoder.decode(parts[1], StandardCharsets.UTF_8);
            } else {
                return "unknown-file";
            }
        } catch (Exception e) {
            return "error-filename";
        }
    }







}
