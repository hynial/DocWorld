package me.util;

import me.caseuse.weekreportbuild.entiry.CoordinateEntity;
import org.apache.poi.hssf.model.InternalSheet;
import org.apache.poi.hssf.record.DVRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.aggregates.DataValidityTable;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;

public class ExcelUtil {
    // 行高度默认保留原高度，有些不行就配置下面两种方式：1、自适应Auto Fit Height；2计算新的高度值
//    public static List<CoordinateEntity> KeepOriginalHeightCoors = Arrays.asList(new CoordinateEntity(24, 1)); // 保留原始高度 // 默认

    // 计算高度 配置
    public static List<CoordinateEntity> ComputeHeightCoors = Arrays.asList(new CoordinateEntity(7, 1));        // 计算固定高度
    // 自适应高度 配置
    public static List<CoordinateEntity> AutoFitHeightCoors = Arrays.asList(new CoordinateEntity(10, 2), new CoordinateEntity(11, 2), new CoordinateEntity(12, 2),
            new CoordinateEntity(13, 2), new CoordinateEntity(14, 2), new CoordinateEntity(15, 2), new CoordinateEntity(16, 2),
            new CoordinateEntity(19, 2), new CoordinateEntity(20, 2), new CoordinateEntity(21, 2), new CoordinateEntity(22, 2));

    public static void fillVariables(XSSFWorkbook destWorkbook, Sheet destinationSheet, Map<String, List<CoordinateEntity>> varMap, Map<String, String> varValueMap) {
        List<String> allVars = varMap.keySet().stream().collect(Collectors.toList());
        List<String> actualVars = varValueMap.keySet().stream().collect(Collectors.toList());

        if (actualVars.size() < allVars.size()) {
            for (String k : allVars) {
                if (!actualVars.contains(k)) {
                    varValueMap.put(k, "");  // 默认值
                }
            }
        }

//        XSSFCellStyle style = targetWorkbook.createCellStyle();
//        style.setWrapText(true);

        Workbook targetWorkbook = destWorkbook == null ? destinationSheet.getWorkbook() : destWorkbook;
        for (Map.Entry<String, String> entry : varValueMap.entrySet()) {
            String var = entry.getKey();
            String val = entry.getValue();

            List<CoordinateEntity> coordinateEntities = varMap.get(var);
            if (coordinateEntities == null || coordinateEntities.isEmpty()) {
                System.out.println("EmptyCoordinate:" + var);
                continue;
            }
            // if (true) continue;
            for (CoordinateEntity coordinateEntity : coordinateEntities) {
                int row = coordinateEntity.row - 1;
                int col = coordinateEntity.col - 1;

                Row rowLine = destinationSheet.getRow(row);
                Cell cell;
                if (rowLine == null) {
                    rowLine = destinationSheet.createRow(row);
                    cell = rowLine.createCell(col);
                    cell.setCellType(CellType.STRING);
                } else {
                    cell = rowLine.getCell(col);
                    if (cell == null) {
                        cell = rowLine.createCell(col);
                        cell.setCellType(CellType.STRING);
                    }
                }

                String originalValue = cell.getStringCellValue();
                if (originalValue == null) {
                    originalValue = "";
                }

                String resultValue = originalValue.replaceAll(var, val);
                // replace \t to 6 space
                resultValue = resultValue.replaceAll("\t", "      ");
                cell.setCellValue(resultValue);

                // set row height
                CellStyle rowStyle = rowLine.getRowStyle();
                if (rowStyle == null) {
                    rowStyle = targetWorkbook.createCellStyle();
                }
                rowStyle.setWrapText(true);
                CellStyle cellStyle = cell.getCellStyle();
                if (cellStyle == null) {
                    cellStyle = targetWorkbook.createCellStyle();
                }
                cellStyle.setWrapText(true);
            }
        }

        // 调整行高度 - 计算
        for (CoordinateEntity coordinateEntity : ComputeHeightCoors) { // 自计算高度
            Row rowLine = destinationSheet.getRow(coordinateEntity.row - 1);
            Cell cell = rowLine.getCell(coordinateEntity.col - 1);
            int numberOfLines = cell.getStringCellValue().split("\n").length;
            rowLine.setHeightInPoints(numberOfLines * destinationSheet.getDefaultRowHeightInPoints());
        }
        // 调整行高度 - 自适应
        for (CoordinateEntity coordinateEntity : AutoFitHeightCoors) {
            Row rowLine = destinationSheet.getRow(coordinateEntity.row - 1);
            rowLine.setHeight((short) -1); // 自适应
        }
    }


    private static boolean equalsRegion(final CellRangeAddressList region1, final CellRangeAddressList region2) {
        return equalsSqref(convertSqref(region1), convertSqref(region2));
    }

    private static boolean equalsSqref(final List<String> sqref1, final List<String> sqref2) {
        if (sqref1.size() != sqref2.size()) return false;

        Collections.sort(sqref1);
        Collections.sort(sqref2);

        final int size = sqref1.size();
        for (int i = 0; i < size; ++i) {
            if (!sqref1.get(i).equals(sqref2.get(i))) {
                return false;
            }
        }

        return true;
    }


    private static List<String> convertSqref(final CellRangeAddressList region) {
        List<String> sqref = new ArrayList<String>();
        for (CellRangeAddress range : region.getCellRangeAddresses()) {
            sqref.add(range.formatAsString());
        }

        return sqref;
    }

    // 同表更新操作 oldRegion 的数据验证 更新到 newRegion
    public static boolean updateDataValidationRegion(final Sheet sheet, final CellRangeAddressList oldRegion, final CellRangeAddressList newRegion) {
        if (sheet instanceof XSSFSheet) {
            List<String> oldSqref = convertSqref(oldRegion);

            try {
                final XSSFSheet xssfSheet = (XSSFSheet) sheet;
                Field fWorksheet = XSSFSheet.class.getDeclaredField("worksheet");
                fWorksheet.setAccessible(true);
                CTWorksheet worksheet = (CTWorksheet) fWorksheet.get(xssfSheet);

                CTDataValidations dataValidations = worksheet.getDataValidations();
                if (dataValidations == null) return false;

                for (int i = 0; i < dataValidations.getCount(); ++i) {
                    CTDataValidation dv = dataValidations.getDataValidationArray(i);

                    @SuppressWarnings("unchecked") List<String> sqref = new ArrayList<String>(dv.getSqref());
                    if (equalsSqref(sqref, oldSqref)) {
                        List<String> newSqref = convertSqref(newRegion);
                        dv.setSqref(newSqref);
                        dataValidations.setDataValidationArray(i, dv);
                        return true;
                    }

                }

                return false;
            } catch (Exception e) {
                throw new RuntimeException("fail update DataValidation's Region.", e);
            }
        } else if (sheet instanceof HSSFSheet) {
            final HSSFSheet hssfSheet = (HSSFSheet) sheet;
            try {
                Field fWorksheet = HSSFSheet.class.getDeclaredField("_sheet");
                fWorksheet.setAccessible(true);
                InternalSheet worksheet = (InternalSheet) fWorksheet.get(hssfSheet);

                DataValidityTable dvt = worksheet.getOrCreateDataValidityTable();

                final AtomicBoolean updated = new AtomicBoolean(false);
                org.apache.poi.hssf.record.aggregates.RecordAggregate.RecordVisitor visitor = new org.apache.poi.hssf.record.aggregates.RecordAggregate.RecordVisitor() {

                    @Override
                    public void visitRecord(Record r) {
                        if (!(r instanceof DVRecord)) return;

                        final DVRecord dvRecord = (DVRecord) r;
                        final CellRangeAddressList region = dvRecord.getCellRangeAddress();
                        if (equalsRegion(region, oldRegion)) {

                            while (region.countRanges() != 0) region.remove(0);

                            for (CellRangeAddress newRange : newRegion.getCellRangeAddresses())
                                region.addCellRangeAddress(newRange);

                            updated.set(true);
                        }
                    }
                };

                dvt.visitContainedRecords(visitor);

                return updated.get();

            } catch (Exception e) {
                throw new RuntimeException("fail update DataValidation's Region.", e);
            }
        }
        throw new UnsupportedOperationException("not supported update dava validation's region for type " + sheet.getClass().getName());
    }

    // 复制数据验证，支持不通sheet表格
    public static boolean copyDataValidationRegion(final Sheet sheet, final Sheet anotherSheet, final CellRangeAddressList oldRegion, final CellRangeAddressList newRegion) {
        if (sheet instanceof XSSFSheet) {
            List<String> oldSqref = convertSqref(oldRegion);

            try {
                final XSSFSheet xssfSheet = (XSSFSheet) sheet;
                Field fWorksheet = XSSFSheet.class.getDeclaredField("worksheet");
                fWorksheet.setAccessible(true);
                CTWorksheet worksheet = (CTWorksheet) fWorksheet.get(xssfSheet);

                CTDataValidations dataValidations = worksheet.getDataValidations();
                List<DataValidation> dataValidationList = new ArrayList<>();
                List<XSSFDataValidation> oldDataValidationList = xssfSheet.getDataValidations();

                if (dataValidations == null) return false;

                oldDataValidationList.forEach(o -> dataValidationList.add(o));

                for (int i = 0; i < dataValidations.getCount(); ++i) {
                    CTDataValidation dv = dataValidations.getDataValidationArray(i);    // 获取区域Square
                    XSSFDataValidation dataV = oldDataValidationList.get(i);            // 获取约束

                    @SuppressWarnings("unchecked") List<String> sqref = new ArrayList<String>(dv.getSqref());
                    if (equalsSqref(sqref, oldSqref)) {
                        XSSFSheet an = (XSSFSheet) anotherSheet;
                        XSSFDataValidationHelper validationHelper = new XSSFDataValidationHelper(an);
                        DataValidation validation = validationHelper.createValidation(dataV.getValidationConstraint(), newRegion);
                        an.addValidationData(validation); // add validation
                        return true;
                    }
                }

                return false;

            } catch (Exception e) {
                throw new RuntimeException("fail update DataValidation's Region.", e);
            }
        }

        throw new UnsupportedOperationException("not supported update dava validation's region for type " + sheet.getClass().getName());
    }
}