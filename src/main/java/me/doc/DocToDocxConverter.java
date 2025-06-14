//package me.doc;
//import org.apache.poi.hwpf.HWPFDocument;
//import org.apache.poi.hwpf.converter.DocxConverter;
//import org.apache.poi.hwpf.usermodel.Range;
//
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.OutputStream;
//
//public class DocToDocxConverter {
//
//    public void convertDocToDocx(String docFilePath, String docxFilePath) throws Exception {
//        try (FileInputStream fis = new FileInputStream(docFilePath);
//             HWPFDocument doc = new HWPFDocument(fis);
//             FileOutputStream fos = new FileOutputStream(docxFilePath)) {
//
//            Range range = doc.getRange();
//            DocxConverter converter = new DocxConverter();
//            converter.process(doc, fos); // 注意：这里应该传递HWPFDocumentCore类型，但HWPFDocument需要适配或转换
//            // 注意：由于HWPFDocument和DocxConverter的API可能有所变化，上述代码可能需要调整。
//            // 一种解决方案是使用Apache POI提供的兼容层或查找最新的API文档来正确实现转换。
//
//            // 另一种更简单的解决方案（如果不需要复杂的格式转换）是：
//            // 1. 读取doc文件内容
//            // 2. 创建一个新的docx文件，并将内容写入其中
//            // 但这种方法可能会丢失一些格式信息。
//        }
//    }
//
//    public static void main(String[] args) {
//        String docFilePath = "sample.doc";
//        String docxFilePath = "sample.docx";
//        DocToDocxConverter converter = new DocToDocxConverter();
//        try {
//            converter.convertDocToDocx(docFilePath, docxFilePath);
//            System.out.println("Conversion completed successfully.");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//}