package me.caseuse.weekreportbuild;

import me.caseuse.weekreportbuild.entiry.CoordinateEntity;
import me.caseuse.weekreportbuild.util.WeekReportUtil;
import me.doc.excel.PoiCopySheet;
import me.util.ExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class WeekReportMain2 {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        String currentDir = System.getProperty("user.dir");
        // Variables
        String varPath = currentDir + File.separator +  "0Input" + File.separator + "excelCoordinate.txt";
        Map<String, List<CoordinateEntity>> varMap = WeekReportUtil.loadAllVarCoordinate(varPath);
        List<String> vars = varMap.keySet().stream().collect(Collectors.toList());

        // Contents
        List<String> contentLines = Files.readAllLines(Paths.get(currentDir, "0Input", "report", "WeekReport250613.txt"), StandardCharsets.UTF_8);
        Map<String, String> varValueMap = WeekReportUtil.getTemplateValueMap(vars, contentLines);

        // Output
        String outDir = "/Users/hynial/PreArchiveCorps/RSZH/0.工作周报";
        // Backup
        String backUp = outDir + File.separator + "backup";
        String tmpDir = outDir + File.separator + "tmp";
        File createDir = new File(backUp);
        if (!createDir.exists()) {
            createDir.mkdirs();
        }
        createDir = new File(tmpDir);
        if (!createDir.exists()) {
            createDir.mkdirs();
        }

        Path originalReport = Files.list(Paths.get(outDir)).filter(path -> path.toFile().getName().indexOf("陈燕辉") > -1 && !path.toFile().getName().contains("updating")).findFirst().get();
        File originalReportFile = originalReport.toFile();
        String replaceText = originalReportFile.getName().substring(originalReportFile.getName().length() - 5);
        Files.copy(originalReport, Paths.get(backUp, originalReportFile.getName()), StandardCopyOption.REPLACE_EXISTING);
        Path updateReport = Paths.get(originalReportFile.getParent(), originalReportFile.getName().replaceFirst(replaceText, "updating" + replaceText));
        Files.copy(originalReport, updateReport, StandardCopyOption.REPLACE_EXISTING);

        String reportName = "安硕-陈燕辉-工作周报（20250606）.xlsx";
        String todayFormat = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyyMMdd"));
        String dateFormat = "";  // Custom Report Date
        if (dateFormat.equals("")) {
            dateFormat = todayFormat;
        }
        reportName = reportName.replaceFirst("\\d+", dateFormat);

        String templatePath = "/Users/hynial/PreArchiveCorps/RSZH/0.工作周报/工作周报2.xlsx";
        File file = new File(templatePath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        File updateReportFile = updateReport.toFile();
        XSSFWorkbook targetWorkbook = new XSSFWorkbook(updateReportFile);

        // Get your sourceSheet object, your code may be different
        Sheet sourceSheet = workbook.getSheetAt(0);

        // Create your destination sheet object, you may have a different sheet name
        String newSheetName = dateFormat;
        int idx = targetWorkbook.getSheetIndex(newSheetName);
        if (idx > -1) {
            targetWorkbook.removeSheetAt(idx);
        }
        Sheet destinationSheet = targetWorkbook.createSheet(newSheetName);

        // Perform the copy using PoiCopySheet.copySheet()
        PoiCopySheet.copySheet(sourceSheet, destinationSheet);

        // replace vars
        // destinationSheet.getRow(12).getCell(1).setCellValue("Changed");
        ExcelUtil.fillVariables(targetWorkbook, destinationSheet, varMap, varValueMap);

        String targetPath = tmpDir + File.separator + reportName;
        File exportFile = new File(targetPath);
        if (exportFile.exists()) {
            exportFile.delete();
        }
        workbookWriteTo(targetWorkbook, targetPath); // 导出新的excel会更新编辑的excel，不导出就不会 。。。。。

        // delete original, replace it with new.
        String originalName = originalReportFile.getName();
        if (originalReportFile.exists()) {
            originalReportFile.delete();
        }

        Files.move(updateReport, updateReport.resolveSibling(reportName));
    }

    private static void cloneSheet(String reportPath, String targetPath) throws IOException, InvalidFormatException {
        File file = new File(reportPath);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        workbook.cloneSheet(0);
        workbook.setSheetName(1, "Cloned");


        try (FileOutputStream outputStream = new FileOutputStream(targetPath)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("Workbook saved to " + targetPath);
    }

    private static void workbookWriteTo(XSSFWorkbook workbook, String targetPath) {
        try (FileOutputStream outputStream = new FileOutputStream(targetPath)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("workbookWriteTo:" + targetPath);
    }
}
