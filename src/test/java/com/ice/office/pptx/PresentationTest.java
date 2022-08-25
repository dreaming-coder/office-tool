package com.ice.office.pptx;

import org.apache.poi.xslf.usermodel.XSLFChart;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;

import java.util.List;

@DisplayName("Presentation方法测试")
public class PresentationTest {
    private Presentation presentation;

    @Test
    @DisplayName("测试读取PPTX文件")
    @BeforeEach
    void testOpenPresentation(){
        presentation = new Presentation("test.pptx");
    }

    @Test
    @DisplayName("测试输出文件")
    void testWritePresentation(){
        presentation.writeTo("output.pptx");
    }

    @Test
    @DisplayName("测试获取演示文稿所有图表")
    void testGetAllChartsOfPresentation(){
        List<XSLFChart> charts = presentation.getCharts();
        // getCTChart() 是获取xml格式的chart
        // 后面是xml文件的标签获取
        List<CTBarChart> barChartList = charts.get(0).getCTChart().getPlotArea().getBarChartList();
        System.out.println(charts.size());
    }

    @Test
    @DisplayName("测试从指定幻灯片获取图表")
    void testGetChartOfSlide(){
        XSLFChart chart = presentation.ChartAt(0);

    }


    @Test
    void test(){
        List<XSLFChart> charts = presentation.getCharts();
        XSLFChart chart = charts.get(0);
        System.out.println(chart.getParent() == presentation.getSlides().get(0));
    }
}
