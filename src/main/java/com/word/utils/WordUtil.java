package com.word.utils;


import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.compress.utils.Lists;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.List;

@Slf4j
public class WordUtil {
    /**
     * 段落
     */
    private static final String P_0 = "本次检测风险项分布如下图:";
    private static final String P_1 = "风险项分布图:";
    private static final String P_2 = "本次检测风险项检测占比如下图:";
    private static final String P_3 = "风险项检测占比图:";
    private static final String P_4 = "加固壳识别:";
    private static final String P_5 = "第三方SDK检测:";
    /**
     * 表格
     */
    private static final String TB_0 = "SDK名称";
    private static final String TB_1 = "权限";
    private static final String TB_2 = "加固壳识别";
    private static final String TB_3 = "权限详情";
    private static final String TB_4 = "风险详情";
 
    public static void main(String[] args) throws Exception {
        press("掌上工程", "/Users/xugan/person/doc/2023-2024八年地理寒假作业解析（四).docx","/Users/xugan/person/a.docx");
    }
 
 
    public static void press(String appName, String filePath,String outPath) throws Exception {
 
        // 延迟解析比率
        ZipSecureFile.setMinInflateRatio(-1.0d);
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        CTBody body = doc.getDocument().getBody();
        body.getPArray(5).getRArray(0).getTArray(0).setStringValue("中国移动自有APP安全评估报告");
        body.getPArray(6).getRArray(0).getTArray(0).setStringValue(appName);
        // 删除图片
        doc.getParagraphs().forEach(WordUtil::removeParagraphPicture);
 
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        XWPFHeader headers = headerFooterPolicy.getDefaultHeader();
        // 获取页眉
        List<XWPFHeader> headerList = doc.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            // 清除页眉页脚，并清除水印
            xwpfHeader.clearHeaderFooter();
        }
        XWPFParagraph paragraph = headers.createParagraph();
        // 页眉左对齐
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        // 页眉底部边框
        paragraph.setBorderBottom(Borders.HEART_GRAY);
        XWPFRun run = paragraph.createRun();
        // 设置页眉
        run.setText("hutools ");
 
        // 移除第一个图片控件，
        body.removeSdt(0);
        // 表格里的第三方SDK
        List<String> sdkList = Lists.newArrayList();
        // 需要移除的表格下标
        List<Integer> tbIndex = Lists.newArrayList();
        List<XWPFTable> tables = doc.getTables();
        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            if (i == 1) {
                // 往基本信息表格加一行
                addRow(doc, table);
            } else {
                // 其他表格删除风险详情
                delRow(table);
            }
            // 表格第一行
            XWPFTableRow row = table.getRow(0);
            // 行列数
            int size = row.getTableCells().size();
            // 表格行超过一列
            if (size > 1) {
                // 第一列文本值
                String text0 = row.getCell(0).getText();
                // 第二列文本值
                String text1 = row.getCell(1).getText();
                if (TB_0.equals(text0)) {
                    System.out.println(text0);
                    tbIndex.add(i);
                    sdkList.add(text1);
                }
                if (TB_1.equals(text0) && size == 5) {
                    System.out.println(text0);
                    tbIndex.add(i);
                }
                if (TB_2.equals(text1)) {
                    tbIndex.add(i);
                }
            } else {
                String text = row.getCell(0).getText();
                if (TB_3.equals(text)) {
                    System.out.println("0\t" + text);
                    tbIndex.add(i);
                }
            }
        }
        // 段落值清空 (第一页)
        body.getPArray(14).getRArray(0).getTArray(0).setStringValue(null);
        // 段落值清空 (第一页)
        body.getPArray(15).getRArray(0).getTArray(0).setStringValue(null);
        // 移除的段落
        List<Integer> removePs = Lists.newArrayList();
        // 所有段落
        List<CTP> pList = body.getPList();
        for (int i = 0; i < pList.size(); i++) {
            List<CTR> rList = pList.get(i).getRList();
            for (CTR ctr : rList) {
                if (ctr.sizeOfTArray() > 0) {
                    String value = ctr.getTArray(0).getStringValue();
                    if (StringUtils.contains(value, P_0) || StringUtils.contains(value, P_1)
                            || StringUtils.contains(value, P_2) || StringUtils.contains(value, P_3)
                    ) {
                        removePs.add(i);
                        System.out.println(value + "\t" + i);
                        break;
                    } else if (StringUtils.contains(value, P_4) || StringUtils.contains(value, P_5)) {
                        removePs.add(i);
                        System.out.println(value + "\t" + i);
                        break;
                    } else if (CollectionUtils.containsAny(sdkList, StringUtils.substringAfter(value, "."))) {
                        removePs.add(i);
                        System.out.println("sdk++++++++++:" + value + i);
                        break;
                    }
                }
            }
        }
 
        // 移除不需要的段落
        int p = 0;
        for (int i = 0; i < removePs.size(); i++) {
            if (i == 0) {
                body.removeP(removePs.get(i));
            } else {
                body.removeP(removePs.get(i) - (1 + p));
                p++;
            }
        }
        // 移除不需要的表格
        int t = 0;
        for (Integer index : tbIndex) {
            if (index == 3) {
                body.removeTbl(index);
            } else {
                body.removeTbl(index - (1 + t));
                t++;
            }
        }
        //输出文件
        FileOutputStream os = new FileOutputStream(outPath);
        doc.write(os);
        doc.close();
        os.close();
        System.out.println("输出的路径："+outPath);
//        System.out.println("文件名为："+name);
    }
 
    /**
     * 移除图片
     */
    public static void removeParagraphPicture(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        try {
            for (XWPFRun run : runs) {
                List<XWPFPicture> pics = run.getEmbeddedPictures();
                if (!pics.isEmpty()) {
                    paragraph.removeRun(0);
                }
            }
        } catch (Exception e) {
            log.error("removeParagraphPicture error...", e);
        }
    }
 
    /**
     * 基础信息表格加一栏 是否加固
     */
    private static void addRow(XWPFDocument doc, XWPFTable table) {
        // 获取加固识别
        String text = doc.getTableArray(3).getRow(3).getCell(1).getText();
        // 表格加一行
        XWPFTableRow row = table.createRow();
        XWPFTableCell cell = row.getCell(0);
        XWPFRun xwpfRun = cell.getParagraphArray(0).createRun();
        xwpfRun.setBold(true);
        xwpfRun.setText("是否加固");
        cell.setColor(table.getRow(0).getCell(0).getColor());
        row.getCell(1).setText("安全".equals(text) ? "是" : "否");
    }
 
    /**
     * 移除表格 风险详情
     */
    private static void delRow(XWPFTable xwpfTable) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (int j = 0; j < rows.size(); j++) {
            XWPFTableRow tableRow = rows.get(j);
            if (StringUtils.contains(tableRow.getCell(0).getText(), TB_4)) {
                xwpfTable.removeRow(j);
                break;
            }
        }
    }
 
}