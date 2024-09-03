package com.word.controller;

import cn.hutool.core.util.StrUtil;
import com.alibaba.fastjson2.JSON;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.SEPX;
import org.apache.poi.hwpf.model.SectionTable;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@RestController
@RequestMapping("/api")
public class FileUploadController {

    @PostMapping("/upload")
    public void handleFileUpload(@RequestParam("files") MultipartFile[] files, HttpServletResponse response) {
        response.setContentType("application/zip");
        response.setHeader("Content-Disposition", "attachment; filename=\"files.zip\"");

        try (
             ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())) {

            for (MultipartFile file : files) {
                if (file.isEmpty()) continue;

                String fileName = file.getOriginalFilename();
                try (ByteArrayInputStream bais = new ByteArrayInputStream(file.getBytes())) {
                    if (fileName != null && fileName.endsWith(".docx")) {
                        processDocxFile(bais, zipOut, fileName);
                    } else if (fileName != null && fileName.endsWith(".doc")) {
                        processDocFile(file.getInputStream(), zipOut, fileName);
                    }
                }
            }

        } catch (Exception e) {
            log.error("", e);
        }
    }

    private void processDocxFile(ByteArrayInputStream bais, ZipOutputStream zipOut, String fileName) throws IOException {

        try (XWPFDocument document = new XWPFDocument(bais)) {
            // Modify document (for demonstration purposes)
//            List<XWPFHeader> headerList = document.getHeaderList();
//            for (XWPFHeader xwpfHeader : headerList) {
//                xwpfHeader.clearHeaderFooter();
//            }
            document.getHeaderList().forEach(XWPFHeader::clearHeaderFooter);
            document.getFooterList().forEach(XWPFFooter::clearHeaderFooter);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            document.write(baos);

            ZipEntry zipEntry = new ZipEntry(fileName);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(baos.toByteArray());
            zipOut.closeEntry();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void processDocFile(ByteArrayInputStream bais, ZipOutputStream zipOut, String fileName) throws IOException {
        try (HWPFDocument document = new HWPFDocument(bais);) {
            // 获取所有页眉
//            HeaderStories headerStories = new HeaderStories(document);

//            // 清除所有页眉内容
//            clearHeader(headerStories.getFirstHeaderSubrange()); // 清除首页页眉
//            clearHeader(headerStories.getOddHeaderSubrange()); // 清除奇数页页眉
//            clearHeader(headerStories.getEvenHeaderSubrange()); // 清除偶数页页眉
//
//            // 清除所有页眉内容
//            clearHeader(headerStories.getFirstFooterSubrange()); // 清除首页页眉
//            clearHeader(headerStories.getOddFooterSubrange()); // 清除奇数页页眉
//            clearHeader(headerStories.getEvenFooterSubrange()); // 清除偶数页页眉

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            document.write(baos);

            ZipEntry zipEntry = new ZipEntry(fileName);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(baos.toByteArray());
            zipOut.closeEntry();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void processDocFile(InputStream inputStream, ZipOutputStream zipOut, String fileName) throws IOException {
        try (HWPFDocument document = new HWPFDocument(inputStream);
             ByteArrayOutputStream baos = new ByteArrayOutputStream()) {

            document.write(baos);

            ZipEntry zipEntry = new ZipEntry(fileName);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(baos.toByteArray());
            zipOut.closeEntry();
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    private static void clearHeader(Range range) {
        if (range != null) {
            String text = range.text();
            if (StrUtil.isNotEmpty(text)) {
                range.replaceText(text, ""); // 清除页眉内容
            }
        }
    }

}
