package com.ice.office.xlsx;

import com.ice.office.domain.xlsx.Sheet;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;
import java.util.Map;

public class SheetTest extends WorkbookTest{
    protected static Sheet sheet;

    @BeforeAll
    static void getSheet(){
        sheet = workbook.getSheetAt(0);
    }

    @Test
    @DisplayName("测试获取series名称")
    void testGetSeriesArray(){
        List<Map<String, String>> seriesArray = sheet.getSeriesArray();
        System.out.println(seriesArray);
    }

    @Test
    @DisplayName("测试获取categories名称")
    void testGetCategoriesArray(){
        List<Map<String, String>> categoriesArray = sheet.getCategoriesArray();
        System.out.println(categoriesArray);
    }

    @Test
    @DisplayName("测试获取表格值称")
    void testGetSerValList(){
        List<List<Map<String, String>>> serValList = sheet.getSerValList();
        serValList.forEach(list->list.forEach(System.out::println));
    }

    @Test
    void test(){
        System.out.println(sheet.getSheet().getRow(1).getCell(1).getReference());
        System.out.println(sheet.getSheet().getRow(1).getCell(1).getRowIndex());
    }
}
