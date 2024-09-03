package com.word.controller;

import cn.hutool.core.util.StrUtil;
import com.alibaba.fastjson2.JSON;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.SEPX;
import org.apache.poi.hwpf.model.SectionTable;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
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
        try {
            // 设置下载的 ZIP 文件名和响应头
            String zipFileName = "processed_docs.zip";
            response.setContentType("application/zip");
            response.setHeader("Content-Disposition", "attachment; filename=\"" + zipFileName + "\"");

            // 创建 ZIP 输出流
            try (ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())) {
                for (MultipartFile file : files) {
                    String originalFileName = file.getOriginalFilename();
                    if (originalFileName == null) continue;

                    // 根据文件扩展名选择处理方法
                    if (originalFileName.endsWith(".doc")) {
                        processDocFile(file.getInputStream(), zipOut, originalFileName);
                    } else if (originalFileName.endsWith(".docx")) {
                        processDocxFile(file.getInputStream(), zipOut, originalFileName);
                    }
                }
            }
        } catch (Exception e) {
            log.error("", e);
        }
    }

    private void processDocxFile(InputStream bais, ZipOutputStream zipOut, String fileName) throws IOException {
        try (XWPFDocument document = new XWPFDocument(bais)) {

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

    // 处理 .doc 文件
    private void processDocFile(InputStream inputStream, ZipOutputStream zipOut, String fileName) throws IOException {
        try (HWPFDocument document = new HWPFDocument(inputStream)) {
            removeHeadersAndFooters(document);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            document.write(baos);

            // 添加到 ZIP 文件
            ZipEntry zipEntry = new ZipEntry(fileName);
            zipOut.putNextEntry(zipEntry);
            zipOut.write(baos.toByteArray());
            zipOut.closeEntry();
        }
    }
    // 去除 .doc 文件的页眉和页脚
    private void removeHeadersAndFooters(HWPFDocument document) {
        HeaderStories headerStories = new HeaderStories(document);

        // 清除首页页眉
        clearHeaderOrFooter(headerStories.getFirstHeaderSubrange());

        // 清除奇数页页眉
        clearHeaderOrFooter(headerStories.getOddHeaderSubrange());

        // 清除偶数页页眉
        clearHeaderOrFooter(headerStories.getEvenHeaderSubrange());

        // 清除首页页脚
        clearHeaderOrFooter(headerStories.getFirstFooterSubrange());

        // 清除奇数页页脚
        clearHeaderOrFooter(headerStories.getOddFooterSubrange());

        // 清除偶数页页脚
        clearHeaderOrFooter(headerStories.getEvenFooterSubrange());
    }

    private void clearHeaderOrFooter(Range range) {
        if (range != null) {
            // 遍历段落
            for (int i = 0; i < range.numParagraphs(); i++) {
                Paragraph paragraph = range.getParagraph(i);
                // 找到段落中所有图片的标识符，并删除它们
                for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
                    String text = paragraph.getCharacterRun(j).text();
                    paragraph.replaceText(text, "", text.length());
                }
            }
        }
    }



}
