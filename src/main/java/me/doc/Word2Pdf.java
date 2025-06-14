package me.doc;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.spire.doc.FileFormat;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Word2Pdf {

    public static void main(String[] args) {
//        convert();
        vv();
    }


    // itext + poi ，没有常识成功，各种版本匹配不上，最后因为缺少谷歌字体失败 GAE。
    public static void convert() {
        try {
            String template = "docground/房贷：担保意向书(农行).docx";
            FileInputStream fileInputStream = new FileInputStream(template);
//            XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
//            PdfOptions pdfOptions = PdfOptions.create();
//            FileOutputStream fileOutputStream = new FileOutputStream("docground/output/房贷：担保意向书(农行)1.docx");
//            PdfConverter.getInstance().convert(xwpfDocument,fileOutputStream,pdfOptions);
//            fileInputStream.close();
//            fileOutputStream.close();


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    // spire-doc 是可用，不过对 pdf 的页数大小有限制
    public static void v() {
        String template = "docground/房贷：担保意向书(农行).docx";
        long start = System.nanoTime();
        //实例化Document类的对象
        com.spire.doc.Document doc = new com.spire.doc.Document();

        // 下载Word文件
//        URL url = new URL("http://xxxx/ExportWord_230724_172956.docx");
//        InputStream inputStream = url.openStream();
//        doc.loadFromStream(inputStream,FileFormat.Docx);
        //加载Word
        doc.loadFromFile(template);
        //保存为PDF格式"docground/output/房贷：担保意向书(农行)1.docx");

        doc.saveToFile("docground/output/房贷：担保意向书(农行)1.pdf", FileFormat.PDF);
        long end = System.nanoTime();
        System.out.println(end - start);
    }

    // aspose 相对推荐使用，不过在 linux 上可能产生中文乱码，来自于缺少字体库
    // aspose-diagram 还可以用于 vsdx 转 pdf
    // 参考：https://blog.csdn.net/weixin_38409915/article/details/125317664
    public static void vv() {
        FileOutputStream os = null;
        try {
            String outPath = "docground/output/房贷：担保意向书(农行)2.pdf";
            String inPath = "docground/房贷：担保意向书(农行).docx";
            long old = System.currentTimeMillis();
            File file = new File(outPath); // 新建一个空白pdf文档
            os = new FileOutputStream(file);
            Document doc = new Document(inPath); // Address是将要被转化的word文档
            doc.save(os, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
            // EPUB, XPS, SWF 相互转换
            long now = System.currentTimeMillis();
            System.out.println("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if (os != null) {
                try {
                    os.flush();
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
