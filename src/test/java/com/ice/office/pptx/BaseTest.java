package com.ice.office.pptx;

import com.ice.office.domain.pptx.Presentation;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.*;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;


public class BaseTest {
    protected static Presentation presentation;

    @DisplayName("读取PPTX文件")
    @BeforeAll
    static void testOpenPresentation() {
        presentation = new Presentation("test.pptx");
    }


    @DisplayName("输出文件")
    @AfterAll
    static void testWritePresentation() {
        presentation.writeTo("output.pptx");
    }
}
