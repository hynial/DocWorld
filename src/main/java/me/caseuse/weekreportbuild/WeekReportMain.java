package me.caseuse.weekreportbuild;

import me.util.DocUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import static me.util.DocUtil.VAR_HEAD;
import static me.util.DocUtil.VAR_TAIL;

public class WeekReportMain {
    private static String TEMPLATE_PATH = "/Users/hynial/PreArchiveCorps/RSZH/金智项目/8.工作汇报/厦门国贸资本集团-陈燕辉_工作周报_Template2.docx";

    private static  Map<String, String> getTemplateValueMap(List<String> vars, List<String> contentLines) {
        Map<String, String> varValueMap = new HashMap<>();
        boolean keyFlag = false;
        String k = "";
        String v = "";
        for (String cLine : contentLines) {
//            String line = cLine.trim();
            String line = cLine;
            if (vars.contains(line)) { // template variable
                if (!keyFlag) {
                    keyFlag = true;
                } else {
                    varValueMap.put(k, v.endsWith("\n") ? v.substring(0, v.length() - 1) : v);
                    v = "";
                }
                k = line;
            } else {
                v += line + "\n";
            }
        }

        if (!varValueMap.containsKey(k)) {
            varValueMap.put(k, v.endsWith("\n") ? v.substring(0, v.length() - 1) : v);
        }

        return varValueMap;
    }

    public static void main(String[] args) throws IOException {
        String currentDir = System.getProperty("user.dir");
        // Variables
        File variablesFile = new File(currentDir + File.separator +  "0Input" + File.separator + "Variables1.txt");
        List<String> varLines = Files.readAllLines(variablesFile.toPath(), StandardCharsets.UTF_8);
        List<String> vars = new ArrayList<>();
        for (String var : varLines) {
            if (var == null) continue;;
            if (var.trim().isEmpty()) continue;
            vars.add(var.trim());
        }

        // Template Content
        List<String> contentLines = Files.readAllLines(Paths.get(currentDir, "0Input", "WeekReport250530.txt"), StandardCharsets.UTF_8);
        Map<String, String> varValueMap = getTemplateValueMap(vars, contentLines);
        String outDir = "/Users/hynial/PreArchiveCorps/RSZH/金智项目/8.工作汇报/周报生成";
        File outDirFile = new File(outDir);
        if (!outDirFile.exists()) {
            outDirFile.mkdirs();
        }

        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
        String newName = sdf.format(new Date());
        try {
            File f = new File(TEMPLATE_PATH);
            InputStream templateStream = Files.newInputStream(f.toPath());
            XWPFDocument document = new XWPFDocument(templateStream);


            boolean keepWeekSix = false;
            boolean keepWeekSeven = false;
            // 删除段落
            int pNumber = document.getParagraphs().size() - 1;
            while (pNumber >= 0) {
                XWPFParagraph p = document.getParagraphs().get(pNumber);
                String pText = p.getParagraphText();

                if (!keepWeekSix && (pText.contains("（周六）") || pText.contains(VAR_HEAD + "SaturdayBody" + VAR_TAIL))) {
                    DocUtil.deleteParagraph(p);
                }
                if (!keepWeekSeven && (pText.contains("（周日）") || pText.contains(VAR_HEAD + "SundayBody" + VAR_TAIL))) {
                    DocUtil.deleteParagraph(p);
                }
                pNumber--;
            }

            DocUtil.replaceVariables(document, varValueMap.keySet().toArray(new String[0]), varValueMap.values().toArray(new String[0]));
            //DocUtil.replaceVariables(document, new String[] {"【Date】"}, new String[]{"20250516"});

            // 将替换后的内容写入新文件
            String targetFilePath = outDir + "/"+ newName +".docx";
            File targetFile = new File(targetFilePath);
            if (targetFile.exists()) {
                targetFile.delete();
            }
            FileOutputStream fos = new FileOutputStream(targetFilePath);
            document.write(fos);
            fos.close();
            templateStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
