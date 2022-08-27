package com.ice.office.xlsx;

import com.ice.office.domain.xlsx.Workbook;
import org.junit.jupiter.api.BeforeAll;

public class BaseTest {

    protected static Workbook workbook;

    @BeforeAll
    static void loadWorkbook(){
        workbook = new Workbook("test.xlsx");
    }
}
