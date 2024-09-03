package com.word.controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.hwpf.HWPFDocument;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.nio.file.*;
import java.util.Comparator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/api2")
public class FileController {

    private static final String UPLOAD_DIR = "/Users/xugan/person/upload/dir"; // 替换为你自己的文件夹路径

    @PostMapping("/upload")
    public void handleFileUpload(@RequestParam("files") MultipartFile[] files, HttpServletResponse response) throws IOException {
        Path uploadPath = Paths.get(UPLOAD_DIR);
        Files.createDirectories(uploadPath);

        // 保存文件到指定文件夹
        for (MultipartFile file : files) {
            Path filePath = uploadPath.resolve(file.getOriginalFilename());
            try (InputStream inputStream = file.getInputStream()) {
                Files.copy(inputStream, filePath, StandardCopyOption.REPLACE_EXISTING);
            }
        }

        // 设置响应头
        response.setContentType("application/zip");
        response.setHeader("Content-Disposition", "attachment; filename=\"processed_files.zip\"");

        try (ZipOutputStream zipOut = new ZipOutputStream(response.getOutputStream())) {
            // 压缩文件夹中的文件
            Files.walk(uploadPath)
                .filter(Files::isRegularFile)
                .forEach(file -> {
                    try (InputStream fis = new FileInputStream(file.toFile())) {
                        ZipEntry zipEntry = new ZipEntry(uploadPath.relativize(file).toString());
                        zipOut.putNextEntry(zipEntry);

                        byte[] buffer = new byte[1024];
                        int length;
                        while ((length = fis.read(buffer)) > 0) {
                            zipOut.write(buffer, 0, length);
                        }
                        zipOut.closeEntry();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                });
            zipOut.finish();
        } catch (IOException e) {
            e.printStackTrace();
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
        } finally {
            // 清理文件夹
            Files.walk(uploadPath)
                .sorted(Comparator.reverseOrder())
                .map(Path::toFile)
                .forEach(File::delete);
        }
    }
}
