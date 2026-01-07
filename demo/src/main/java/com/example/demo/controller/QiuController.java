package com.example.demo.controller;


import com.example.demo.util.ExcelUtil;
import com.example.demo.util.ExcelValue;
import com.google.common.base.Splitter;
import io.micrometer.common.util.StringUtils;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/sjcl")
public class QiuController {
    // 缓存每次生成的Excel字节流和对应文件名
    private List<ExcelData> excelDataList = new ArrayList<>();

    // 内部类：存储Excel字节流和文件名
    private static class ExcelData {
        private byte[] excelBytes; // Excel字节数组
        private String fileName;   // Excel文件名（无后缀）

        public ExcelData(byte[] excelBytes, String fileName) {
            this.excelBytes = excelBytes;
            this.fileName = fileName;
        }
    }
    @PostMapping("1")
    public void qiuSjCl(HttpServletResponse response, @RequestParam("files") List<MultipartFile> files,@RequestParam("file2") MultipartFile file2) throws IOException {
        try {
            XSSFWorkbook xfb = new XSSFWorkbook();
            List<ExcelValue> excelValues = new ArrayList<>();
            for (MultipartFile file : files) {
                InputStream fis = file.getInputStream();
                xfb = new XSSFWorkbook(fis);
                StringJoiner sj = new StringJoiner("-");
                XSSFSheet sheetAt = xfb.getSheetAt(0);
                String sheetName = StringUtils.isEmpty(sheetAt.getSheetName()) ? "原始数据" : sheetAt.getSheetName();
                sj.add(sheetName);
                Iterator<Row> iterator = sheetAt.rowIterator();
                int rowNum = 0;
                Row row = null;

                while (iterator.hasNext()) {
                    row = iterator.next();
                    if (rowNum > 0) {
                        String mx = (row.getCell(1) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(1)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(1))).setScale(2, RoundingMode.DOWN).toString();
                        String power = (row.getCell(2) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(2)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(2))).setScale(2, RoundingMode.DOWN).toString();
                        ExcelValue ev = new ExcelValue();
                        ev.setMx(mx);
                        ev.setPower(power);
                        excelValues.add(ev);
                        String mx1 = (row.getCell(4) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(4)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(4))).setScale(2, RoundingMode.DOWN).toString();
                        String middle = (row.getCell(5) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(5)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(5))).setScale(2, RoundingMode.DOWN).toString();
                        ExcelValue ev1 = new ExcelValue();
                        ev1.setMx(mx1);
                        ev1.setMiddle(middle);
                        excelValues.add(ev1);
                        String mx2 = (row.getCell(7) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(7)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(7))).setScale(2, RoundingMode.DOWN).toString();
                        String thin = (row.getCell(8) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(8)))) ? "" : new BigDecimal(ExcelUtil.getCellValue(row.getCell(8))).setScale(2, RoundingMode.DOWN).toString();
                        ExcelValue ev2 = new ExcelValue();
                        ev2.setMx(mx2);
                        ev2.setThin(thin);
                        excelValues.add(ev2);
                    }
                    rowNum++;
                }
                //取出所有mz值
                List<String> mx = excelValues.stream().map(ExcelValue::getMx).distinct().collect(Collectors.toList());
                //按照所有mz分组
                Map<String, List<ExcelValue>> collect = excelValues.stream().collect(Collectors.groupingBy(ExcelValue::getMx));
                List<ExcelValue> excelValueList = new ArrayList<>();
                for (String s : mx) {
                    List<ExcelValue> values = collect.get(s);
                    ExcelValue ev = new ExcelValue();
                    ev.setMx(s);
                    for (ExcelValue value : values) {
                        String power = value.getPower();
                        String middle = value.getMiddle();
                        String thin = value.getThin();
                        if (!StringUtils.isEmpty(power)) {
                            ev.setPower(power);
                        }
                        if (!StringUtils.isEmpty(middle)) {
                            ev.setMiddle(middle);
                        }
                        if (!StringUtils.isEmpty(thin)) {
                            ev.setThin(thin);
                        }
                    }
                    excelValueList.add(ev);
                }
                excelValueList = processExcelValueList(excelValueList);
                matchDataBaseData(excelValueList, file2, response,file.getOriginalFilename(),files.size());
                excelValueList.clear();
            }
            if(CollectionUtils.isNotEmpty(files)&&files.size()>1){
                exportAllToZip(response,"处理文件");
            }
        } catch(Exception e){
            System.out.println(e);
        }


    }



    //原始数据中如果出现重复数据则对重复数据进行处理取最大的那个
    public static List<ExcelValue> processExcelValueList(List<ExcelValue> excelValueList) {
        // 分组操作，将具有相同 mx 的 ExcelValue 对象分到一组
        Map<String, List<ExcelValue>> groupedMap = excelValueList.stream()
                .collect(Collectors.groupingBy(ExcelValue::getMx));

        List<ExcelValue> resultList = new ArrayList<>();
        for (Map.Entry<String, List<ExcelValue>> entry : groupedMap.entrySet()) {
            String mx = entry.getKey();
            List<ExcelValue> group = entry.getValue();
            if (group.size() > 1) {
                // 当 mx 存在重复项时，找出最大的 power、middle 和 thin
                int maxPower = group.stream().mapToInt(v -> Integer.parseInt(v.getPower())).max().orElse(0);
                int maxMiddle = group.stream().mapToInt(v -> Integer.parseInt(v.getMiddle())).max().orElse(0);
                int maxThin = group.stream().mapToInt(v -> Integer.parseInt(v.getThin())).max().orElse(0);
                // 创建新的 ExcelValue 对象并添加到结果列表
                ExcelValue excelValue = new ExcelValue();
                excelValue.setMx(mx);
                excelValue.setPower(String.valueOf(maxPower));
                excelValue.setMiddle(String.valueOf(maxMiddle));
                excelValue.setThin(String.valueOf(maxThin));
                resultList.add(excelValue);
            } else {
                // 如果 mx 没有重复项，直接将该对象添加到结果列表
                resultList.add(group.get(0));
            }
        }
        return resultList;
    }

    //匹配数据库数据后输出最终文件
    public void matchDataBaseData(List<ExcelValue> excelValues, MultipartFile file, HttpServletResponse response,String fileName,Integer fileCount) throws IOException {
        Map<String, List<ExcelValue>> collect = excelValues.stream().collect(Collectors.groupingBy(ExcelValue::getMx));
        Set<String> strings = collect.keySet();
        InputStream fis = file.getInputStream();
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(fis);
        XSSFSheet sheetAt = xssfWorkbook.getSheetAt(0);
        Iterator<Row> iterator = sheetAt.iterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            System.out.println("数据库匹配中:目前匹配行数" + row.getRowNum());
            if (row.getRowNum() > 0) {
                //数据库需要匹配的mz
                String eValue = "";
                if (row.getCell(4) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(16)))) {
                } else {
                    if (row.getCell(4).getCellTypeEnum().equals(CellType.FORMULA)) {
                        FormulaEvaluator evaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
                        CellValue evaluate = evaluator.evaluate(row.getCell(4));
                        eValue = new BigDecimal(evaluate.getNumberValue()).setScale(2, RoundingMode.DOWN).toString();
                    } else {
                        eValue = new BigDecimal(ExcelUtil.getCellValue(row.getCell(4))).setScale(2, RoundingMode.DOWN).toString();
                    }
                }
                String kValue = "";
                if (row.getCell(10) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(16)))) {
                } else {
                    if (row.getCell(10).getCellTypeEnum().equals(CellType.FORMULA)) {
                        FormulaEvaluator evaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
                        CellValue evaluate = evaluator.evaluate(row.getCell(10));
                        kValue = new BigDecimal(evaluate.getNumberValue()).setScale(2, RoundingMode.DOWN).toString();
                    } else {
                        kValue = new BigDecimal(ExcelUtil.getCellValue(row.getCell(10))).setScale(2, RoundingMode.DOWN).toString();
                    }
                }
                String qValue = "";
                if (row.getCell(16) == null || StringUtils.isEmpty(ExcelUtil.getCellValue(row.getCell(16)))) {
                } else {
                    if (row.getCell(16).getCellTypeEnum().equals(CellType.FORMULA)) {
                        FormulaEvaluator evaluator = xssfWorkbook.getCreationHelper().createFormulaEvaluator();
                        CellValue evaluate = evaluator.evaluate(row.getCell(16));
                        qValue = new BigDecimal(evaluate.getNumberValue()).setScale(2, RoundingMode.DOWN).toString();
                    } else {
                        qValue = new BigDecimal(ExcelUtil.getCellValue(row.getCell(16))).setScale(2, RoundingMode.DOWN).toString();
                    }
                }
                if (strings.contains(eValue)) {
                    checkRowExist(row, 5, 8);
                    List<ExcelValue> values = collect.get(eValue);

                    row.getCell(5).setCellValue(values.get(0).getMx());
                    row.getCell(6).setCellValue(values.get(0).getPower());
                    row.getCell(7).setCellValue(values.get(0).getMiddle());
                    row.getCell(8).setCellValue(values.get(0).getThin());
                }

                if (strings.contains(kValue)) {
                    checkRowExist(row, 11, 14);
                    List<ExcelValue> values = collect.get(kValue);
                    row.getCell(11).setCellValue(values.get(0).getMx());
                    row.getCell(12).setCellValue(values.get(0).getPower());
                    row.getCell(13).setCellValue(values.get(0).getMiddle());
                    row.getCell(14).setCellValue(values.get(0).getThin());
                }

                if (strings.contains(qValue)) {
                    checkRowExist(row, 17, 20);
                    List<ExcelValue> values = collect.get(qValue);
                    row.getCell(17).setCellValue(values.get(0).getMx());
                    row.getCell(18).setCellValue(values.get(0).getPower());
                    row.getCell(19).setCellValue(values.get(0).getMiddle());
                    row.getCell(20).setCellValue(values.get(0).getThin());
                }
            } else {
                this.checkRowExist(row, 5, 8);
                row.getCell(5).setCellValue("m/z(若数据来源源自多个，则按照强中弱由左到右展示来源信息)");
                row.getCell(6).setCellValue("强");
                row.getCell(7).setCellValue("中");
                row.getCell(8).setCellValue("弱");
                this.checkRowExist(row, 11, 14);
                row.getCell(11).setCellValue("m/z(若数据来源源自多个，则按照强中弱由左到右展示来源信息)");
                row.getCell(12).setCellValue("强");
                row.getCell(13).setCellValue("中");
                row.getCell(14).setCellValue("弱");
                this.checkRowExist(row, 17, 20);
                row.getCell(17).setCellValue("m/z(若数据来源源自多个，则按照强中弱由左到右展示来源信息)");
                row.getCell(18).setCellValue("强");
                row.getCell(19).setCellValue("中");
                row.getCell(20).setCellValue("弱");
            }
        }

        if(fileCount>1){
            addExcel(xssfWorkbook,Splitter.on(".").splitToList(fileName).get(0)+"处理文件");
        }else {
            ExportLoadExcel(response, xssfWorkbook, "ManipulateData");
        }
    }

    public Row checkRowExist(Row row, Integer startIndex, Integer endIndex) {
        for (int i = startIndex; i < endIndex + 1; i++) {
            if (row.getCell(i) == null) {
                row.createCell(i);
            }
        }
        return row;
    }
    public void addExcel(XSSFWorkbook xssfWorkbook, String fileName) throws IOException {
        if (xssfWorkbook == null || fileName == null || fileName.trim().isEmpty()) {
            throw new IllegalArgumentException("Excel工作簿和文件名不能为空");
        }

        // 将Excel写入字节数组（缓存到内存）
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            xssfWorkbook.write(baos);
            excelDataList.add(new ExcelData(baos.toByteArray(), fileName.trim()));
        } finally {
            xssfWorkbook.close(); // 释放POI资源
        }
    }

    public void exportAllToZip(HttpServletResponse response, String zipFileName) throws IOException {
        if (excelDataList.isEmpty()) {
            throw new IllegalStateException("未添加任何Excel文件，无法打包ZIP");
        }

        // 1. 设置ZIP响应头
        response.setContentType("application/zip");
        response.setHeader("Access-Control-Allow-Origin", "*");
        response.setCharacterEncoding("utf-8");
        String encodedZipName = URLEncoder.encode(zipFileName, StandardCharsets.UTF_8.name())
                .replaceAll("\\+", "%20");
        response.setHeader("Content-Disposition", "attachment;filename*=utf-8''" + encodedZipName + ".zip");
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");

        // 2. 打包所有缓存的Excel为ZIP并写入响应
        try (OutputStream responseOs = response.getOutputStream();
             ZipOutputStream zipOs = new ZipOutputStream(responseOs)) {

            for (ExcelData excelData : excelDataList) {
                // 创建ZIP条目（每个Excel在ZIP内的名称）
                ZipEntry zipEntry = new ZipEntry(excelData.fileName + ".xlsx");
                zipOs.putNextEntry(zipEntry);
                // 写入Excel字节流
                zipOs.write(excelData.excelBytes);
                zipOs.closeEntry();
            }

            zipOs.flush();
        } finally {
            // 清理缓存，避免重复下载
            excelDataList.clear();
        }
    }

    // ====================== 兼容原有单次导出方法 ======================
    public void exportSingleExcel(HttpServletResponse response, XSSFWorkbook xssfWorkbook, String fileName) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Access-Control-Allow-Origin", "*");
        response.setCharacterEncoding("utf-8");
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8.name())
                .replaceAll("\\+", "%20");
        response.setHeader("Content-Disposition", "attachment;filename*=utf-8''" + encodedFileName + ".xlsx");

        try (OutputStream os = response.getOutputStream()) {
            xssfWorkbook.write(os);
            os.flush();
        } finally {
            xssfWorkbook.close();
        }
    }
    public void ExportLoadExcel(HttpServletResponse response, XSSFWorkbook xssfWorkbook, String fileName) throws IOException {
//        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Access-Control-Allow-Origin", "*");
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename*=utf-8''" + fileName + ".xlsx");
        xssfWorkbook.write(response.getOutputStream());
        response.getOutputStream().close();
    }

}
