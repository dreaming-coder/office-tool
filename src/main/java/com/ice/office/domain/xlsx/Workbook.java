package com.ice.office.domain.xlsx;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

@Slf4j
public class Workbook {

    // 工作簿
    private XSSFWorkbook workbook;

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public Workbook() {
    }

    public Workbook(String path) {
        try {
            this.workbook = new XSSFWorkbook(Files.newInputStream(Paths.get(path)));
            log.info("打开文件{}成功", path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取指定的sheet页
     *
     * @param index sheet页索引
     * @return 指定的sheet页对象
     */
    public Sheet getSheetAt(int index) {
        return new Sheet(workbook.getSheetAt(index));
    }

    /**
     * 获取指定的sheet页
     *
     * @param name sheet页名称
     * @return 指定的sheet页对象
     */
    public Sheet getSheet(String name) {
        return new Sheet(workbook.getSheet(name));
    }

}
