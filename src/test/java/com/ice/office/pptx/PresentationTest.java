package com.ice.office.pptx;

import com.ice.office.domain.pptx.Presentation;
import com.ice.office.domain.pptx.Slide;
import com.ice.office.domain.pptx.chart.Chart;
import org.junit.jupiter.api.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.util.List;

@DisplayName("Presentation方法测试")
public class PresentationTest extends BaseTest{

    @Test
    @DisplayName("测试获取演示文稿所有图表")
    void testGetSlides(){
        List<Slide> slides = presentation.getSlides();
        System.out.println(slides.size());
    }

    @Test
    @DisplayName("测试获取演示文稿所有图表")
    void testGetCharts(){
        List<Chart> charts = presentation.getCharts();
        // getCTChart() 是获取xml格式的chart
        // 后面是xml文件的标签获取
        List<CTBarChart> barChartList = charts.get(0).getCtChart().getPlotArea().getBarChartList();
        System.out.println(charts.size());
        System.out.println(barChartList.size());
    }

    @Test
    @DisplayName("测试从指定幻灯片获取图表")
    void testGetChartOfSlide(){
        Chart chart = presentation.chartAt(0);
        System.out.println(chart);
    }

    @Test
    @DisplayName("测试从指定幻灯片获取图表XML标签")
    void testGetCTChartOfSlide(){
        CTChart ctChart = presentation.ctChartAt(0);
        System.out.println(ctChart);
    }
}
