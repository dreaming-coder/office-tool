package com.ice.office.pptx;

import com.ice.office.domain.pptx.Slide;
import org.junit.jupiter.api.BeforeAll;

public class SlideTest extends BaseTest{
    protected static Slide slide;

    @BeforeAll
    static void getSlide(){
        slide = presentation.getSlides().get(0);
    }
}
