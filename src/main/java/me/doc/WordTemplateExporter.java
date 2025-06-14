package me.doc;

//import org.apache.poi.hwpf.HWPFDocument;
//import org.apache.poi.hwpf.usermodel.CharacterRun;
//import org.apache.poi.hwpf.usermodel.Paragraph;
//import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

public class WordTemplateExporter {
    public static void main(String[] args) {
        try {
            String template = "docground/房贷：担保意向书(农行).docx";
            // 从 resources 中加载模板文件
            // InputStream templateStream = WordTemplateExporter.class.getResourceAsStream("docground/房贷：担保意向书(农行).doc");
            InputStream templateStream = Files.newInputStream(Paths.get(template));

//            HWPFDocument document = new HWPFDocument(templateStream);
//            replacePlaceholder(document, "[#放款银行#]", "中国人民银行厦门分行");


            XWPFDocument document = new XWPFDocument(templateStream);
            replacePlaceholder(document, "[#放款银行#]", "中国人民银行厦门分行");

            // 将替换后的内容写入新文件
            FileOutputStream fos = new FileOutputStream("docground/output/1.docx");
            document.write(fos);
            fos.close();
            templateStream.close();
            System.out.println("Word 文件导出成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

//    private static void replacePlaceholder(HWPFDocument document, String placeholder, String replacement) {
//        Range range = document.getRange();
//        for (int i = 0; i < range.numParagraphs(); i++) {
//            Paragraph paragraph = range.getParagraph(i);
//            for (int j = 0; j < paragraph.numCharacterRuns(); j++) {
//                CharacterRun characterRun = paragraph.getCharacterRun(j);
//                String text = characterRun.text();
//                int index = text.indexOf(placeholder);
//                System.out.println(index);
//                characterRun.replaceText(placeholder, replacement);
//            }
//        }
//    }

    private static void replacePlaceholder(XWPFDocument document, String placeholder, String replacement) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                if (text != null && text.contains(placeholder)) {
                    text = text.replace(placeholder, replacement);
                    run.setText(text, 0);
                }
            }
        }
    }

    public static boolean replaceAllText(String docPath, String[] toFindText, String[] newText) {
        if (toFindText == null || newText == null) {
            return true;
        }

        if (toFindText.length != newText.length) {
            throw new IllegalArgumentException("查找值数量与替换值数量不等");
        }

        try {
            InputStream templateStream = Files.newInputStream(Paths.get(docPath));
            XWPFDocument document = new XWPFDocument(templateStream);

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);

                    for(int i = 0; i < toFindText.length; i++) {
                        String placeholder = toFindText[i];
                        String replacement = newText[i];
                        if (text != null && text.contains(placeholder)) {
                            text = text.replace(placeholder, replacement);
                            run.setText(text, 0);
                        }
                    }
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return true;
    }
}