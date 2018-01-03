package com.soul;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Array;
import java.util.*;

import org.apache.poi.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.NilPICFAndBinData;
import org.apache.poi.hwpf.model.TextPiece;
import org.apache.poi.hwpf.model.TextPieceTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

class WordRead {

    private HWPFDocument hwpfDocument;

    private XWPFDocument xwpfDocument;
    ArrayList<HashMap<String, HashMap<String, String>>> mapArrayList = new ArrayList<HashMap<String, HashMap<String, String>>>();

    WordRead(ArrayList<String> paths) {
        int i = 0;
        while (i < paths.size()) {
            String path = paths.get(i);
            System.out.print(path);
            HashMap<String, HashMap<String, String>> map = new HashMap<String, HashMap<String, String>>();
            if (path.endsWith("doc")) {
                map = readDocWord(path);
            }else if (path.endsWith("docx")) {
                map = readDocxWord(path);
            }
            mapArrayList.add(map);
            i++;
        }

        for (HashMap<String, HashMap<String, String>> map: mapArrayList) {
            System.out.println(map);
        }
    }

    private HashMap<String, HashMap<String, String>> readDocxWord(String path) {
        File file = new File(path);
        HashMap<String, HashMap<String, String>> map = new HashMap<String, HashMap<String, String>>();

        HashMap<String, String> userMap = new HashMap<String, String>();
        HashMap<String, String> resultMap = new HashMap<String, String>();
        XWPFTable resultTable = null;

        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            xwpfDocument = new XWPFDocument(fileInputStream);
            List<XWPFTable> tableList = xwpfDocument.getTables();

            for (XWPFTable table : tableList) {
                ArrayList<String> resultList = new ArrayList<String>();
                for (int i = 0; i < table.getNumberOfRows(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        XWPFTableCell cell = row.getCell(j);
                        for (int k = 0; k < cell.getParagraphs().size(); k++) {
                            XWPFParagraph paragraph = cell.getParagraphArray(k);
                            String s = paragraph.getText();
                            System.out.println(s);
                            if (s.contains("报告时间")) { return map; }
                            System.out.println(s);
                            String str = s.trim();
                            if (str.contains("：")) {
                                String[] array = str.split("：");
                                if (array.length >= 2) {
                                    String key = array[0];
                                    String value = array[1];
                                    userMap.put(key, value);
                                }
                            }else if (str.contains(":")) {
                                String[] array = str.split(":");
                                if (array.length >= 2) {
                                    String key = array[0];
                                    String value = array[1];
                                    userMap.put(key, value);
                                }
                            }else {
                                if (str.contains("检测项目")) {
                                    resultTable = table;
                                }
                                if (table.equals(resultTable)) {
                                    System.out.println("result ===" + str);
                                    if (!str.contains("检测项目") && !str.contains("检测结果")) {
                                        resultList.add(str);
                                    }
                                }
                            }
                        }
                    }
                }
                if (resultList.size() > 0) {
                    for (int i = 0; i < resultList.size(); i++) {
                        if (i%2 == 0) {
                            if (i+1>= resultList.size()) { break; }
                            String key = resultList.get(i);
                            String value = resultList.get(i+1);
                            resultMap.put(key,value);
                        }
                    }
                }
                map.put("user", userMap);
                map.put("result", resultMap);
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private HashMap<String, HashMap<String, String>> readDocWord(String path) {
        File file = new File(path);
        HashMap<String, HashMap<String, String>> map = new HashMap<String, HashMap<String, String>>();

        HashMap<String, String> userMap = new HashMap<String, String>();
        HashMap<String, String> resultMap = new HashMap<String, String>();

        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            hwpfDocument = new HWPFDocument(fileInputStream);

            Range range = hwpfDocument.getRange();
            TableIterator iterator = new TableIterator(range);

            Table resultTable = null;

            while (iterator.hasNext()) {
                Table table = (Table) iterator.next();
                ArrayList<String> resultList = new ArrayList<String>();
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow tableRow = table.getRow(i);
                    for (int j = 0; j < tableRow.numCells(); j++) {
                        TableCell tableCell = tableRow.getCell(j);// 取得单元格
                        for (int k = 0; k < tableCell.numParagraphs(); k++) {
                            Paragraph paragraph = tableCell.getParagraph(k);
                            String s = paragraph.text();
                            if (s.contains("报告时间")) { return map; }
                            System.out.println(s);
                            String str = s.trim();
                            if (str.contains("：")) {
                                String[] array = str.split("：");
                                if (array.length >= 2) {
                                    String key = array[0];
                                    String value = array[1];
                                    userMap.put(key, value);
                                }
                            }else if (str.contains(":")) {
                                String[] array = str.split(":");
                                if (array.length >= 2) {
                                    String key = array[0];
                                    String value = array[1];
                                    userMap.put(key, value);
                                }
                            }else {
                                if (str.contains("检测项目")) {
                                    resultTable = table;
                                }
                                if (table.equals(resultTable)) {
                                    System.out.println("result ===" + str);
                                    if (!str.contains("检测项目") && !str.contains("检测结果")) {
                                        resultList.add(str);
                                    }
                                }
                            }
                        }
                    }
                }
                if (resultList.size() > 0) {
                    for (int i = 0; i < resultList.size(); i++) {
                        if (i%2 == 0) {
                            if (i+1>= resultList.size()) { break; }
                            String key = resultList.get(i);
                            String value = resultList.get(i+1);
                            resultMap.put(key,value);
                        }
                    }
                }
                map.put("user", userMap);
                map.put("result", resultMap);
            }
        }catch (Exception e) {
            e.printStackTrace();
        }

        return map;
    }

}



