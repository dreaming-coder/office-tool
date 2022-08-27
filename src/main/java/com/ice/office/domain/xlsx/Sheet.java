package com.ice.office.domain.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.*;

public class Sheet {
    private XSSFSheet sheet;
    private int rowNUm;
    private int colNum;

    private String sheetName;

    public String getSheetName() {
        return sheetName;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public Sheet() {
    }

    public Sheet(XSSFSheet sheet) {
        this.sheet = sheet;
        this.rowNUm = sheet.getLastRowNum() + 1;
        this.colNum = sheet.getRow(0).getLastCellNum();
        this.sheetName = sheet.getSheetName();
    }

    /**
     * 获取表格系列值及其对应的单元格地址
     *
     * @return 系列值字符串数组
     */
    public List<Map<String, String>> getSeriesArray() {
        List<Map<String, String>> series = new LinkedList<>();
        Map<String, String> map;
        for (Cell cell : sheet.getRow(0)) {
            map = new HashMap<>();
            map.put("address", formatAddress(cell.getAddress().formatAsString()));
            map.put("value", cell.getStringCellValue());
            series.add(map);
        }
        return series;
    }

    /**
     * 返回表格类型值及其单元格地址
     *
     * @return 类型值字符串数组
     */
    public List<Map<String, String>> getCategoriesArray() {
        List<Map<String, String>> categories = new LinkedList<>();
        Map<String, String> map;
        for (int i = 1; i < rowNUm; i++) {
            map = new HashMap<>();
            map.put("address", formatAddress(sheet.getRow(i).getCell(0).getReference()));
            map.put("value", sheet.getRow(i).getCell(0).getStringCellValue());
            categories.add(map);
        }
        return categories;
    }

    /**
     * 返回表格具体内容，以List存储，每一个元素对应一个series的所有值及其对应的单元格地址
     *
     * @return 表格内容
     */
    public List<List<Map<String, String>>> getSerValList() {
        List<List<Map<String, String>>> list = new LinkedList<>();

        for (int col = 1; col < colNum; col++) {
            List<Map<String, String>> seriesVal = new LinkedList<>();
            for (int row = 1; row < rowNUm; row++) {
                Map<String, String> map = new HashMap<>();
                map.put("address", formatAddress(sheet.getRow(row).getCell(col).getReference()));
                map.put("value", sheet.getRow(row).getCell(col).toString());
                seriesVal.add(map);
            }
            list.add(seriesVal);
        }
        return list;
    }

    private String formatAddress(String address) {
        return "$" + address.charAt(0) + "$" + address.charAt(1);
    }
}
