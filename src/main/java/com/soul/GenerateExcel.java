package com.soul;

import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.record.FontRecord;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
/*
* 两种excel 风格
* Quantitative 定量报告
* Screening 筛查报告
*/
enum ExcelStyle {
    Quantitative,
    Screening
}

class GenerateExcel {

    private ArrayList<HashMap<String, HashMap<String, String>>> mapArrayList;

    private static HSSFWorkbook workbook = new HSSFWorkbook();
    private static HSSFSheet sheet_1 = workbook.getSheet("定量结果");
    private static HSSFSheet sheet_2 = workbook.getSheet("筛查结果");

    /*
    * 定量表的 病人信息 表头
    * */
    private ArrayList<String> user_info_keys_1 = new ArrayList<String>();
    /*
    * 筛查表的 病人信息 表头
    * */
    private ArrayList<String> user_info_keys_2 = new ArrayList<String>();

    private ArrayList<String> resultKeys_1 = new ArrayList<String>();

    private ArrayList<String> resultKeys_2 = new ArrayList<String>();

    GenerateExcel(ArrayList<HashMap<String, HashMap<String, String>>> mapArrayList) {

        try {
            FileInputStream inp = new FileInputStream("result.xls");
            workbook = new HSSFWorkbook(inp);

        }catch (Exception e) {
            e.printStackTrace();
        }
        this.mapArrayList = mapArrayList;
        sheet_1 = workbook.getSheet("定量结果");
        sheet_2 = workbook.getSheet("筛查结果");
        if (sheet_1 == null) {
            sheet_1 = workbook.createSheet("定量结果");
        }
        if (sheet_2 == null) {
            sheet_2 = workbook.createSheet("筛查结果");
        }

        user_info_keys_1.addAll(Arrays.asList("姓名", "性别", "年龄", "标本编号", "送检医院", "科室", "床号", "送检医生", "住院号", "标本类型", "临床诊断", "采集日期", "送检日期"));

        user_info_keys_2.addAll(Arrays.asList("姓名", "性别", "年龄", "标本编号", "病区", "床号", "住院号",  "临床诊断", "标本来源"));


        resultKeys_1.addAll(Arrays.asList("CBFβ/MYH11融合基因", "CBFβ/MYH11(拷贝数)", "ABL(拷贝数)", "CBFβ/MYH11/ABL"));

        resultKeys_2.addAll(Arrays.asList("MLL/AF6", "MLL/AF9", "MLL/AF10", "MLL/AF17", "MLL/ELL", "dupMLL", "AML1/ETO", "PML/RARa", "PLZF/RARa", "NPM/RARa", "CBFB/MYH11", "NPM/MLF1", "TLS/ERG", "DEK/CAN", "内部对照", "阴性对照"));
    }

    void generate() {
        try {
            FileOutputStream fileOut = new FileOutputStream("result.xls");

            // 先计算当前文件的row个数
            System.out.println(sheet_1.getLastRowNum());
            System.out.println(sheet_2.getLastRowNum());

            createTableHeader(fileOut);
            for (HashMap<String, HashMap<String, String>> map: mapArrayList) {
                HashMap<String, String> user = map.get("user");
                HashMap<String, String> result = map.get("result");

                CCRow ccRow = new CCRow(user, result);
                ccRow.addRow();
            }
            CreationHelper creationHelper = workbook.getCreationHelper();
            try {
                workbook.write(fileOut);
            }catch (Exception e) {
                e.printStackTrace();
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
    }

    private class CCRow {
        ExcelStyle style;
        HashMap<String, String> user;
        HashMap<String, String> result;


        CCRow(HashMap<String, String> user, HashMap<String, String> result) {
            this.result = result;
            this.user = user;
            if (result.size() <= 5) {
                this.style = ExcelStyle.Quantitative;
            }else  {
                this.style = ExcelStyle.Screening;
            }
        }

        private void addResultWith(HSSFSheet sheet) {
            HSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);

            row.setHeight((short) 600);
            HSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            HSSFFont font = workbook.createFont();
            font.setFontHeight((short) 240);
            cellStyle.setFont(font);

            ArrayList<String> user_info_keys = style == ExcelStyle.Quantitative ? user_info_keys_1 : user_info_keys_2;
            ArrayList<String> resultKeys = style == ExcelStyle.Quantitative ? resultKeys_1: resultKeys_2;

            for (int i = 0; i < user_info_keys.size(); i++) {
                String key = user_info_keys.get(i);
                String value = user.get(key);
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(value);
                cell.setCellStyle(cellStyle);
            }
            System.out.println(result);
            for (int j = 0; j < resultKeys.size(); j++) {
                int index = user_info_keys.size() + j;
                sheet.autoSizeColumn(index);
                String key = resultKeys.get(j);
                String value = result.get(key);
                HSSFCell cell = row.createCell(index);
                cell.setCellValue(value);
                cell.setCellStyle(cellStyle);
            }
        }

        void addRow() {
            if (style == ExcelStyle.Quantitative) {
                addResultWith(sheet_1);
            }else {
                addResultWith(sheet_2);
            }
        }
    }


    /*
    * 创建表头
    * */
    private void createTableHeader(FileOutputStream fileOut) {

        tableHeaderWith(user_info_keys_1, resultKeys_1, sheet_1);
        tableHeaderWith(user_info_keys_2, resultKeys_2, sheet_2);

    }


    private void tableHeaderWith(ArrayList<String> user_info_keys, ArrayList<String> resultKeys, HSSFSheet sheet) {
        if (sheet.getLastRowNum() > 0) {
            return;
        }

        HSSFRow row = sheet.createRow(0);
        sheet.setAutobreaks(true);
        sheet.autoSizeColumn(0);
        row.setHeight((short) 600);
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setFontHeight((short) 240);
        cellStyle.setFont(font);

        for (int i = 0; i < user_info_keys.size(); i++) {
            sheet.addMergedRegion(new CellRangeAddress(0,1, i, i));
            String cellText = user_info_keys.get(i);
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(cellText);
            cell.setCellStyle(cellStyle);
            sheet.autoSizeColumn(i);
        }

        HSSFCell cell_1 = row.createCell(user_info_keys.size());
        cell_1.setCellStyle(cellStyle);
        if (sheet.equals(sheet_1)) {
            cell_1.setCellValue("定量结果");
        }else {
            cell_1.setCellValue("筛查结果");
        }

        HSSFRow row2 = sheet.createRow(1);
        row2.setHeight((short) 600);

        sheet.addMergedRegion(new CellRangeAddress(0,0, user_info_keys.size(), user_info_keys.size()+resultKeys.size()-1));
        for (int j = user_info_keys.size(); j < user_info_keys.size()+resultKeys.size(); j++) {
            int index = j - user_info_keys.size();
            sheet.autoSizeColumn(j);
            String cellText = resultKeys.get(index);
            HSSFCell cell = row2.createCell(j);
            cell.setCellValue(cellText);
            cell.setCellStyle(cellStyle);
        }

    }


}
