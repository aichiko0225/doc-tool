package com.soul;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

import org.apache.poi.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.TextPiece;
import org.apache.poi.hwpf.model.TextPieceTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;

class WordRead {

    private HWPFDocument hwpfDocument;

    private XWPFDocument xwpfDocument;

    WordRead(ArrayList<String> paths) {
        int i = 0;
        ArrayList<HashMap<String, String>> mapArrayList = new ArrayList<HashMap<String, String>>();
        while (i < paths.size()) {
            String path = paths.get(i);
            System.out.print(path);
            HashMap<String, String> map = new HashMap<String, String>();
            if (path.endsWith("doc")) {
                map = readDocWord(path);
            }else if (path.endsWith("docx")) {
                map = readDocxWord(path);
            }
            mapArrayList.add(map);
            i++;
        }

        for (HashMap<String, String> map: mapArrayList) {
            System.out.println(map);
        }
    }

    private HashMap<String, String> readDocxWord(String path) {

        File file = new File(path);
        Map<String, String> map = new HashMap<String, String>();

        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            xwpfDocument = new XWPFDocument(fileInputStream);
            List<XWPFTable> tableList = xwpfDocument.getTables();

            for (XWPFTable table : tableList) {
                for (int i = 0; i < table.getNumberOfRows(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        XWPFTableCell cell = row.getCell(j);
                        for (int k = 0; k < cell.getParagraphs().size(); k++) {
                            XWPFParagraph paragraph = cell.getParagraphArray(k);
                            String s = paragraph.getText();
                            System.out.println(s);
                        }
                    }
                }
            }

        }catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }


    private HashMap<String, String> readDocWord(String path) {
        File file = new File(path);
        Map<String, String> map = new HashMap<String, String>();

        try {
            FileInputStream fileInputStream = new FileInputStream(file);
            hwpfDocument = new HWPFDocument(fileInputStream);

            Range range = hwpfDocument.getRange();
            TableIterator iterator = new TableIterator(range);

            while (iterator.hasNext()) {
                Table table = (Table) iterator.next();
                for (int i = 0; i < table.numRows(); i++) {
                    TableRow tableRow = table.getRow(i);
                    for (int j = 0; j < tableRow.numCells(); j++) {
                        TableCell tableCell = tableRow.getCell(j);// 取得单元格
                        for (int k = 0; k < tableCell.numParagraphs(); k++) {
                            Paragraph paragraph = tableCell.getParagraph(k);
                            String s = paragraph.text();
                            System.out.println(s);
                        }
                    }
                }
            }
        }catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}



