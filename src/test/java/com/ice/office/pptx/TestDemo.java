package com.ice.office.pptx;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;

@Slf4j
public class TestDemo {
    public static void main(String[] args) {
        try {
            XMLSlideShow xmlSlideShow = new XMLSlideShow(Files.newInputStream(new File("test.pptx").toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    @BeforeEach
    public void s(){
        System.out.println("hello");
        log.debug("debug message");
        log.info("info message");
        log.warn("warn message");
        log.error("error message");
    }
}
