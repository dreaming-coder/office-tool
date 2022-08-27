package com.ice.office.xlsx;

import com.ice.office.domain.xlsx.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.junit.jupiter.api.Test;

public class WorkbookTest extends BaseTest{


    @Test
    void testGetSheet(){
        Sheet sheet = workbook.getSheetAt(0);
        Sheet sheet1 = workbook.getSheet("Sheet1");
        System.out.println(sheet == sheet1);
    }

    @Test
    void testGetSeriesArray(){

    }
}
