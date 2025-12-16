package com.example.demo.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;

/**
 * @description ${内容}
 * @date 2021/9/10 10:08
 * @author: lxs
 */
public class ExcelUtil {

    public static void setReportStyle(Workbook workbook, Cell cell, short color, short fh, String fontName) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//左右居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//上下居中
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderRight(BorderStyle.THIN);// 右边框
        cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(156, 195, 230)));//设置背景色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充模式
        Font font = workbook.createFont();
        font.setColor(color);
        font.setFontHeightInPoints(fh);
        font.setFontName(fontName);
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);
    }

    public static void setHeadRow(XSSFWorkbook wb, XSSFSheet sheet, String planYear, List<Map<String, String>> masterInfo, XSSFRow headRow, String roleType,String loginDeptCode) {
        XSSFCell cell = null;
        headRow.setHeight((short) 700);
        for (int i = 0; i < 4; i++) {

            if (i == 0) {
                cell = headRow.createCell(i);
                cell.setCellStyle(bgColor(wb));
                cell.setCellValue(roleType.contains("FZB")?"fzb":"zyb");
            }if (i == 1) {
                cell = headRow.createCell(i);
                cell.setCellStyle(bgColor(wb));
                cell.setCellValue(loginDeptCode+","+planYear);
            }if (i == 2) {
                cell = headRow.createCell(i);
                cell.setCellStyle(bgColor(wb));
                cell.setCellValue("单位");
                sheet.setColumnWidth(i, 9500);
            }if (i > 2) {
                for (int j = 0; j < masterInfo.size(); j++) {
                    cell = headRow.createCell(j*3+4);
                    cell.setCellStyle(bgColor(wb));
                    cell = headRow.createCell(j*3+5);
                    cell.setCellStyle(bgColor(wb));
                    cell = headRow.createCell(j*3+3);
                    cell.setCellStyle(bgColor(wb));

                    cell.setCellValue(masterInfo.get(j).get("masterTargetName")+" ("+ masterInfo.get(j).get("measUnit")+")");
                }
            }
        }
    }

    public static void setHiddenRow(String planYear, String speciality, XSSFRow hiddenRow,List<Map<String, String>> masterInfo,List<String> fields) {

        XSSFCell cell = null;
        for (int i = 0; i < 4; i++) {

            if (i == 0) {
                cell = hiddenRow.createCell(i);
                cell.setCellValue("deptCode");
                fields.add("deptCode");
            }if (i == 1) {
                cell = hiddenRow.createCell(i);
                cell.setCellValue("deptClass");
                fields.add("deptClass");
            }if (i == 2) {
                cell = hiddenRow.createCell(i);
                cell.setCellValue("deptName");
                fields.add("deptName");
            }
            if (i > 2) {
                for (int j = 0; j < masterInfo.size(); j++) {
                    cell = hiddenRow.createCell(j*3+3);
                    cell.setCellValue(masterInfo.get(j).get("masterTargetCode"));
                    fields.add(masterInfo.get(j).get("masterTargetCode"));
                }
            }
        }
    }

    public static XSSFCellStyle unLockStyle(XSSFWorkbook workbook) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setLocked(false);
        return cellStyle;
    }

    /**
     * 获取单元格样式
     * @param workbook
     * @return
     */
    public static XSSFCellStyle bgColor(Workbook workbook) {
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(239, 226, 149)));
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);//填充模式
        cellStyle.setAlignment(HorizontalAlignment.CENTER);//左右居中
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);//上下居中
        cellStyle.setBorderBottom(BorderStyle.THIN); // 下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);// 左边框
        cellStyle.setBorderTop(BorderStyle.THIN);// 上边框
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 合并单元格
     * @param sheet
     * @param masterInfo
     */
    public static void setMergedRegion(XSSFSheet sheet, List<Map<String, String>> masterInfo) {
        //标题行
        CellRangeAddress region = new CellRangeAddress(0, 0, 2, masterInfo.size()*3+2);
        sheet.addMergedRegion(region);
        for (int i = 1; i <= masterInfo.size(); i++) {
            //第二行
            CellRangeAddress region1 = new CellRangeAddress(1, 1, 3*i, 3*i+2);
            sheet.addMergedRegion(region1);
            //第三行
            region=new CellRangeAddress(2, 2, 3*i, 3*i+2);
            sheet.addMergedRegion(region);
        }
        //第四行
        region=new CellRangeAddress(2, 3, 2, 2);
        sheet.addMergedRegion(region);
    }

    /**
     * 解析格式
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        String cellValue = "";
        // 以下是判断数据的类型
        switch (cell.getCellTypeEnum()) {
            case NUMERIC: // 数字
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    cellValue = sdf.format(org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell.getNumericCellValue())).toString();
                } else {
                    DataFormatter dataFormatter = new DataFormatter();
                    cellValue = dataFormatter.formatCellValue(cell);
                }
                break;
            case STRING: // 字符串
                cellValue = cell.getStringCellValue();
                break;
            case BOOLEAN: // Boolean
                cellValue = cell.getBooleanCellValue() + "";
                break;
            case FORMULA: // 公式
                cellValue = cell.getCellFormula() + "";
                break;
            case BLANK: // 空值
                cellValue = "";
                break;
            case ERROR: // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }
}
