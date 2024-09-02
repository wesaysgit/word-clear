package com.word.controller;

import cn.hutool.core.util.StrUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.HeaderStories;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
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
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
@RestController
@RequestMapping("/api")
public class FileUploadController {

    @PostMapping("/upload")
    public ResponseEntity<byte[]> handleFileUpload(@RequestParam("files") MultipartFile[] files) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream();
             ZipOutputStream zipOut = new ZipOutputStream(baos)) {

            for (MultipartFile file : files) {
                if (file.isEmpty()) continue;

                String fileName = file.getOriginalFilename();
                try (ByteArrayInputStream bais = new ByteArrayInputStream(file.getBytes())) {
                    if (fileName != null && fileName.endsWith(".docx")) {
                        processDocxFile(bais, zipOut, fileName);
                    } else if (fileName != null && fileName.endsWith(".doc")) {
                        processDocFile(bais, zipOut, fileName);
                    }
                }
            }

            zipOut.finish();

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=modified-files.zip");
            headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

            return new ResponseEntity<>(baos.toByteArray(), headers, HttpStatus.OK);
        } catch (Exception e) {
            log.error("", e);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
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
            // 获取所有图片数据
            // 获取所有图片数据
            List<XWPFPictureData> pictures = document.getAllPictures();

            if (!pictures.isEmpty()) {
                // 获取最后一张图片
                XWPFPictureData lastPicture = pictures.get(pictures.size() - 1);

                // 遍历文档中的段落
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    List<XWPFRun> runs = paragraph.getRuns();
                    if (runs != null) {
                        for (XWPFRun run : runs) {
                            // 检查并移除图片
                            removePictureFromRun(run, lastPicture);
                        }
                    }
                }

                // 移除图片数据
                document.getPackage().getParts().remove(lastPicture.getPackagePart());

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

    private void processDocFile(ByteArrayInputStream bais, ZipOutputStream zipOut, String fileName) throws IOException {
        try (HWPFDocument document = new HWPFDocument(bais);) {
            // 获取所有页眉
            HeaderStories headerStories = new HeaderStories(document);

            // 清除所有页眉内容
            clearHeader(headerStories.getFirstHeaderSubrange()); // 清除首页页眉
            clearHeader(headerStories.getOddHeaderSubrange()); // 清除奇数页页眉
            clearHeader(headerStories.getEvenHeaderSubrange()); // 清除偶数页页眉

            // 清除所有页眉内容
            clearHeader(headerStories.getFirstFooterSubrange()); // 清除首页页眉
            clearHeader(headerStories.getOddFooterSubrange()); // 清除奇数页页眉
            clearHeader(headerStories.getEvenFooterSubrange()); // 清除偶数页页眉

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

    private static void clearHeader(Range range) {
        if (range != null) {
            String text = range.text();
            if (StrUtil.isNotEmpty(text)) {
                range.replaceText(text, ""); // 清除页眉内容
            }
            System.out.println("清除页眉内容");
        }
    }

}
