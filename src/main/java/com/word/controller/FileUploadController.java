package com.word.controller;

import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PAPBinTable;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
//            // 遍历文档中的段落
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            int posId = paragraphs.size() - 1;
////            document.removeBodyElement(posId);
//            if (!paragraphs.isEmpty()) {
//                XWPFParagraph paragraphToRemove = paragraphs.get(posId);
//                int elementPos = document.getPosOfParagraph(paragraphToRemove);
//                document.removeBodyElement(elementPos);
//            }

            boolean imageFound = false;
            // 从后往前遍历段落，查找最后一张图片
//            for (int i = paragraphs.size() - 1; i >= 0 && !imageFound; i--) {
//                XWPFParagraph paragraph = paragraphs.get(i);
//                List<XWPFRun> runs = paragraph.getRuns();
//
//                for (int j = runs.size() - 1; j >= 0 && !imageFound; j--) {
//                    XWPFRun run = runs.get(j);
//                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
//
//                    if (!pictures.isEmpty()) {
//                        paragraph.removeRun(j);  // 移除图片所在的 run
//                        // 检查该段落是否还有其他内容
//                        if (paragraph.getRuns().isEmpty()) {
//                            paragraphs.remove(i);  // 如果段落已空，移除整个段落
//                        }
//                        imageFound = true;
//                    }
//                }
//            }
            XWPFParagraph paragraph = paragraphs.get(posId);
            List<XWPFRun> runs = paragraph.getRuns();
            for (int i = runs.size() - 1; i >= 0 ; i--) {
                XWPFRun run = runs.get(i);
                List<XWPFPicture> pictures = run.getEmbeddedPictures();
                if (!pictures.isEmpty()) {
                    paragraph.removeRun(i);
                }
            }

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

        // 清除 首页/奇数页/偶数页 页眉
        clearHeaderOrFooter(headerStories.getFirstHeaderSubrange());
        clearHeaderOrFooter(headerStories.getOddHeaderSubrange());
        clearHeaderOrFooter(headerStories.getEvenHeaderSubrange());

        // 清除 首页/奇数页/偶数页 页脚
        clearHeaderOrFooter(headerStories.getFirstFooterSubrange());
        clearHeaderOrFooter(headerStories.getOddFooterSubrange());
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
