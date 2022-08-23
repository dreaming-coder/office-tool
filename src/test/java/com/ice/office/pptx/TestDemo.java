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
            new XMLSlideShow(Files.newInputStream(new File("a.pptx").toPath()));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    @BeforeEach
    public void s(){
        System.out.println("hello");
        log.debug("====fsdsfsfsdfsdfds{}",123);
    }
}
