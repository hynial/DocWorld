package me.util;

import org.apache.poi.xwpf.usermodel.*;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocUtil {
    public static final String VAR_HEAD = "∆"; // 这个符号是Option+j键打印
    public static final String VAR_TAIL = "¬"; // 这个符号是Option+l键打印

    private static Pattern pattern = Pattern.compile("\\d+");
    //    private static Pattern var_pattern = Pattern.compile("【[a-zA-Z]+\\d?\\d?】");
    private static Pattern var_pattern = Pattern.compile(VAR_HEAD + "[a-zA-Z]+\\d?\\d?" + VAR_TAIL);
    public static final String REG_NEW_LINE = "\\n";

    public static void main(String[] args) {
        String text = "【aaabdcsd22】【aaabdcsd22】";
        Matcher matcher = var_pattern.matcher(text);
        while (matcher.find()) {
            System.out.println(matcher.group());
        }
    }

    public static boolean replaceAllDigital(XWPFDocument document, String[] toFindText, String[] newText) {
        if (toFindText == null || newText == null) {
            return true;
        }

        if (toFindText.length != newText.length) {
            throw new IllegalArgumentException("查找值数量与替换值数量不等");
        }

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            StringBuilder p = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                p.append(text);
            }

            for(int i = 0; i < toFindText.length; i++) {
                String placeholder = toFindText[i];
                String replacement = newText[i];
                int idx = 0;
                if (p.toString().contains(placeholder)) {
                    for (XWPFRun run : paragraph.getRuns()) {
                        String text = run.getText(0);
                        if (text != null) { //  text.matches(".*\\d.*")
                            Matcher matcher = pattern.matcher(text);

                            while (matcher.find()) {
                                String d = matcher.group();
                                if (d.length() <= 8) {
                                    text = text.replace(d, replacement.substring(idx, idx + d.length()));
                                    idx += d.length();
                                    run.setText(text, 0);
                                }
                            }
                        }
                    }
                }
            }
        }

        return true;
    }

    private static String defaultValue = "";
    public static boolean replaceVariables(XWPFDocument document, String[] variables, String[] values) {
        if (variables == null) {
            return true;
        }

        if (values == null) {
            values = new String[variables.length];
            Arrays.fill(values, defaultValue);
        }

        if (variables.length > values.length) {
            String[] newValues = new String[variables.length];
            for (int i = 0; i < variables.length; i++) {
                if (i < values.length) {
                    newValues[i] = values[i] == null ? defaultValue : values[i];
                } else {
                    newValues[i] = defaultValue;
                }
            }

            values = newValues;
        }

        System.out.println("Vars:" + Arrays.toString(variables) + ", Values:" + Arrays.toString(values));

        // 表格操作
        List<String> variableList = Arrays.asList(variables);
        for(XWPFTable table : document.getTables()) {
            int colIndex = 1; // Index of the column (starting from 0)
            for (int rowIndex = 1; rowIndex < table.getNumberOfRows(); rowIndex++) {
                XWPFTableRow row = table.getRow(rowIndex);
                if (row != null) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    if (colIndex < cells.size()) {
                        XWPFTableCell cell = cells.get(colIndex);
                        String cellValue = cell.getText();
                        if (variableList.contains(cellValue)) {
                            while (cell.getParagraphs().size() > 1) { // 保留一个段落
                                cell.removeParagraph(1);
                            }
                            // cell.setText(values[variableList.indexOf(cellValue)]); // 简单追加文本，没有换行效果

                            XWPFParagraph pg = cell.getParagraphs().get(0);
                            List<XWPFRun> oriRuns = pg.getRuns();
                            for( int ri = oriRuns.size() - 1; ri >= 0; ri--) {
                                pg.removeRun(ri);
                            }
//                            XWPFParagraph pg = document.createParagraph(); // 通过文档创建会在整个文档末尾追加段落，可以通过原来的段落进行替换。
                            String[] lines = values[variableList.indexOf(cellValue)].split(REG_NEW_LINE);
                            XWPFRun newRun = pg.createRun();
                            newRun.setText(lines[0], 0);
                            for (int i = 1; i < lines.length; i++) { // 带有换行效果
                                newRun.addBreak();
                                newRun.setText(lines[i]);
                            }

                            // cell.setParagraph(pg); // 如果是新建的段落，需要把段落覆盖
                        }
                        System.out.println("Value at row " + rowIndex + ", column " + colIndex + ": " + cellValue);
                    } else {
                        System.out.println("Column index out of bounds.");
                    }
                } else {
                    System.out.println("Row index out of bounds.");
                }
            }
        }

        // 段落操作
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            StringBuilder p = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                String text = run.getText(0);
                p.append(text);
            }

            // 段内可能有多个
            Matcher matcher = var_pattern.matcher(p.toString());
            List<String> varVals = new ArrayList<>();
            while (matcher.find()) {
                varVals.add(matcher.group());
            }

            Map<String, Integer> targetVarOfParagraph = new HashMap<>();
            for (int i = 0; i < variables.length; i++) {
                if (varVals.contains(variables[i])) {
                    targetVarOfParagraph.put(variables[i], i);
                }
            }

            if (targetVarOfParagraph.isEmpty()) continue;

            List<XWPFRun> runs = paragraph.getRuns();
            if (p.toString().contains("¬∆")) {
                //System.out.println(); // 调试使用
            }
            // 满足整个模版匹配
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun episode = runs.get(i);
                String episodeText = episode.getText(0);
                if (episodeText == null) {
                    continue;
                }

                String replaceText = episodeText;

                TreeSet<VarSequenceInfo> varSequenceInfoTreeSet = new TreeSet<>(new Comparator<VarSequenceInfo>() {
                    @Override
                    public int compare(VarSequenceInfo o1, VarSequenceInfo o2) {
                        return o1.index - o2.index;
                    }
                });
                for (String var : targetVarOfParagraph.keySet()) {
                    int varIndex = replaceText.indexOf(var);
                    while (varIndex > -1) {
                        varSequenceInfoTreeSet.add(new VarSequenceInfo(var, varIndex, values[targetVarOfParagraph.get(var)]));
                        varIndex = replaceText.indexOf(var, varIndex + var.length());
                    }
                }

                if (!varSequenceInfoTreeSet.isEmpty()) {
                    int fromIndex = 0;
                    int toIndex;
                    int vSize = varSequenceInfoTreeSet.size();
                    for (int v = 0; v < vSize; v++) {
                        VarSequenceInfo varSequenceInfo = varSequenceInfoTreeSet.pollFirst();
                        toIndex = varSequenceInfo.index;
                        String oriText = episodeText.substring(fromIndex, toIndex);
                        String[] ls = varSequenceInfo.val.split(REG_NEW_LINE);
                        if (v == 0) {
                            episode.setText(oriText + ls[0], 0); // 覆盖
                        } else {
                            episode.setText(oriText + ls[0]); // 追加
                        }

                        for (int lineCount = 1; lineCount < ls.length; lineCount++) {
                            episode.addBreak();
                            episode.setText(ls[lineCount]);
                        }

                        fromIndex = toIndex + varSequenceInfo.var.length();
                    }
                }

                // 以下代码不考虑变量值中的换行效果，导致文本错乱
//                for (String var : targetVarOfParagraph.keySet()) {
//                    int varIndex = replaceText.indexOf(var);
//                    if (varIndex > -1)
//                    while (replaceText.contains(var)) {
//                        String[] ls = values[targetVarOfParagraph.get(var)].split(REG_NEW_LINE);
//                        //replaceText = replaceText.replaceAll(var, values[targetVarOfParagraph.get(var)]);
//                        replaceText = replaceText.replace(var, ls[0]);
//                        episode.setText(replaceText, 0);
//                        for (int lineCount = 1; lineCount < ls.length; lineCount++) {
//                            episode.addBreak();
//                            episode.setText(ls[lineCount]);
//                        }
//                    }
//                }

            }

            // 部分模板匹配
            String prevPart = "";
            int prevPartIndex = 0;
            String nextPart = "";
            int nextPartIndex = 0;
            List<PartPatch> partPatches = new ArrayList<>();
            boolean headMatched = false;
            String upStreamLefts = "";
            boolean existUpstream = false;
            String downStreamLefts = "";
            boolean existDownstream = false;
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun episode = runs.get(i);
                String episodeText = episode.getText(0);
                if (episodeText == null) {
                    continue;
                }
                if (existDownstream) {
                    nextPart = downStreamLefts + episodeText;
                    int upstreamIndex = nextPart.indexOf(VAR_HEAD);
                    if (upstreamIndex > -1) {
                        nextPart = nextPart.substring(0, upstreamIndex);
                        upStreamLefts = nextPart.substring(upstreamIndex);
                        existUpstream = true;
                        episode.setText(nextPart, 0);
                    } else {
                        existUpstream = false;
                    }
                    nextPartIndex = i;

                    if (nextPartIndex > prevPartIndex) {
                        String middle = "";
                        for (int t = prevPartIndex + 1; t < nextPartIndex; t++) {
                            middle += runs.get(t).getText(0);
                        }
                        String unitParts = prevPart + middle + nextPart;
                        for (String var : targetVarOfParagraph.keySet()) {
                            if (unitParts.contains(var)) {
                                PartPatch partPatch = new PartPatch(prevPartIndex, nextPartIndex, var, values[targetVarOfParagraph.get(var)]);
                                partPatches.add(partPatch);
                            }
                        }

                    }

                    headMatched = false;
                    existDownstream = false;
                    continue;
                }

                if (existUpstream) {
                    prevPart = upStreamLefts + episodeText;
                    int downStreamIndex = prevPart.indexOf(VAR_TAIL);
                    if (downStreamIndex > - 1) {
                        prevPart = prevPart.substring(0, downStreamIndex);
                        downStreamLefts = prevPart.substring(downStreamIndex);
                        existDownstream = true;
                    } else {
                        existDownstream = false;
                    }
                    prevPartIndex = i;
                    headMatched = true;
                    episode.setText(prevPart, 0);
                    existUpstream = false;
                    continue;
                }

                if (episodeText.contains(VAR_HEAD) && !headMatched) {
                    prevPart = episodeText;
                    prevPartIndex = i;
                    headMatched = true;
                }

                if (episodeText.contains(VAR_TAIL)) {
                    int upstreamIndex = episodeText.indexOf(VAR_HEAD);
                    nextPart = episodeText;
                    if (upstreamIndex > -1) {
                        nextPart = episodeText.substring(0, upstreamIndex);
                        upStreamLefts = episodeText.substring(upstreamIndex);
                        existUpstream = true;
                        episode.setText(nextPart, 0);
                    }
                    nextPartIndex = i;

                    if (nextPartIndex > prevPartIndex) {
                        String middle = "";
                        for (int t = prevPartIndex + 1; t < nextPartIndex; t++) {
                            middle += runs.get(t).getText(0);
                        }
                        String unitParts = prevPart + middle + nextPart;
                        for (String var : targetVarOfParagraph.keySet()) {
                            if (unitParts.contains(var)) {
                                PartPatch partPatch = new PartPatch(prevPartIndex, nextPartIndex, var, values[targetVarOfParagraph.get(var)]);
                                partPatches.add(partPatch);
                            }
                        }

                    }

                    headMatched = false;
                }
            }


            if (!partPatches.isEmpty()) {
                for (PartPatch partPatch : partPatches) {
                    XWPFRun episode = runs.get(partPatch.prevPartIndex);
                    String episodeText = episode.getText(0);

                    int prevIndex = episodeText.indexOf(VAR_HEAD);
                    // 以下两行没有换行效果
//                    String targetText = episodeText.substring(0, prevIndex) + partPatch.val;
//                    episode.setText(targetText, 0);

                    String[] lines = partPatch.val.split(REG_NEW_LINE);
                    String targetText = episodeText.substring(0, prevIndex) + lines[0];
                    episode.setText(targetText, 0); // Set the first line without creating a break
                    for (int i = 1; i < lines.length; i++) {
                        episode.addBreak(); // Add a break before each subsequent line
                        episode.setText(lines[i]);
                    }

                    for (int j = partPatch.prevPartIndex + 1; j < partPatch.nextPartIndex; j++) {
                        episode = runs.get(j);
                        episode.setText("", 0);
                    }

                    episode = runs.get(partPatch.nextPartIndex);
                    episodeText = episode.getText(0);
                    int nextIndex = episodeText.indexOf(VAR_TAIL);
                    targetText = episodeText.substring(nextIndex + 1);
                    episode.setText(targetText, 0);
                }
            }
        }

        return true;
    }

    static class PartPatch {
        public int prevPartIndex;
        public int nextPartIndex;

        public String var;
        public String val;

        public PartPatch(int prevPartIndex, int nextPartIndex, String var, String val) {
            this.prevPartIndex = prevPartIndex;
            this.nextPartIndex = nextPartIndex;
            this.var = var;
            this.val = val;
        }
    }

    static class VarSequenceInfo {
        public String var;
        public String val;
        public int index;

        public VarSequenceInfo(String var, int index, String val) {
            this.var = var;
            this.index = index;
            this.val = val;
        }
    }

    // 删除表格段落用：cell.removeParagraph(cell.getParagraphs().indexOf(para));
    public static void deleteParagraph(XWPFParagraph p) {
        XWPFDocument doc = p.getDocument();
        int pPos = doc.getPosOfParagraph(p);
//        doc.getDocument().getBody().removeP(pPos);
        //p.setPageBreak(false);

        doc.removeBodyElement(pPos);
    }

    // 删除空行-段落
    public static void deleteEmptyParagraph(XWPFDocument document) {
        int pNumber = document.getParagraphs().size() - 1;
        for (int pi = pNumber; pi >= 0; pi--) {
            XWPFParagraph paragraph = document.getParagraphs().get(pi);
            if (paragraph.getText().isEmpty()) {
                document.removeBodyElement(pi);
            }
        }
    }
}
