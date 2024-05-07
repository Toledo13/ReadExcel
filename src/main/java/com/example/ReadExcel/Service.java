package com.example.ReadExcel;


import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;


public class Service {
    private final String url = "C:\\YOUR\\PATH\\DIR";

    Service sv = new Service();

    public String ReadTable() throws IOException {
        ArrayList<String> titles = new ArrayList<>();
        ArrayList<String> content = new ArrayList<>();
        String aux;
        int nColum = sv.ReadSizeColum();
        int nRow = sv.ReadSizeRow();
        int c;
        String fileName = sv.ReadNameExcel();
        if (fileName.contains(".xlsx")) {
            for (int i = 0; i <= nRow; i++) {
                for (int j = 0; j <= nColum; j++) {
                    if (i == 0) {
                        aux = sv.ReadExcelXlsx(fileName, i, j);
                        titles.add(j, aux);
                    } else {
                        aux = sv.ReadExcelXlsx(fileName, i, j);
                        content.add(j, aux);
                    }
                    if (i != 0) {
                        System.out.println(titles.get(j) + ":" + content.get(j));
                    }
                }
            }


        } else {
            System.out.println("FAIL:This file is not XLSX!");
        }

        return null;
    }

    public String ReadNameExcel() throws IOException {
        List<Path> myList = Files.list(Paths.get(url)).collect(Collectors.toList());
        String nameWithDir = String.valueOf(myList.get(0));
        String nameFile = nameWithDir.replace("C:\\YOUR\\PATH\\DIR", "");
        return nameFile;
    }

    public int ReadSizeColum() throws IOException {
        int colum = 0;
        String title;
        Service sv = new Service();
        String nameFile = sv.ReadNameExcel();
        ArrayList<String> list = new ArrayList<>();
        if (nameFile.contains(".xlsx")) {
            for (int i = 0; i <= 100; i++) {
                title = ReadExcelXlsx(nameFile, 0, i);
                list.add(i, title);
                if (!list.get(i).equals("")) {
                } else {
                    i = 100;
                }
            }
        } else {
            System.out.println("FAIL:This file is not XLSX!");
        }
        colum = list.size();
        colum = colum - 1;
        return colum;
    }

    public int ReadSizeRow() throws IOException {
        Service sv = new Service();
        ArrayList<String> cells = new ArrayList<>();
        ArrayList<String> counter = new ArrayList<>();
        String sheet;
        String content;
        int row;
        String nameFile = sv.ReadNameExcel();
        if (nameFile.contains(".xlsx")) {
            for (int i = 0; i <= 10000; i++) {
                for (int j = 0; j < sv.ReadSizeColum(); j++) {
                    content = sv.ReadExcelXlsx(nameFile, i, j);
                    cells.add(j, content);
                    if (j == 2) {
                        if (cells.get(0).equals("") && cells.get(1).equals("")) {
                            i = 10000;
                        }
                    }
                }
                if (i != 10000) {
                    counter.add(i, "Line" + i);
                }
            }
        } else {
            System.out.println("FAIL:This file is not XLSX!");
        }

        row = counter.size();
        row = row - 1;
        return row;
    }

    public String ReadExcelXlsx(String nameFile, int rowNumber, int columNumber) throws IOException {

        String data = "";
        try {
            FileInputStream fis = new FileInputStream(url + nameFile);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            String nameSheet = workbook.getSheetName(0);
            XSSFSheet sheet = workbook.getSheet(nameSheet);
            XSSFRow row = sheet.getRow(rowNumber);
            XSSFCell cell = row.getCell(columNumber);
            try {
                data = cell.getStringCellValue();
            } catch (IllegalStateException e) {
                data = String.valueOf(cell.getNumericCellValue());
            } catch (NullPointerException e) {
                data = "";
            }

        } catch (NullPointerException e) {
            //ignore
        } catch (Exception e) {
            System.out.println("Fail read");
            e.printStackTrace();
        }
        return data;
    }


}
