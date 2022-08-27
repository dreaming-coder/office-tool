package com.ice.office.pptx.chart;

import com.ice.office.domain.pptx.chart.BarChart;
import com.ice.office.domain.pptx.chart.Chart;
import com.ice.office.domain.xlsx.Workbook;
import com.ice.office.pptx.BaseTest;
import org.apache.poi.xddf.usermodel.chart.XDDFTitle;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.List;

public class ChartTest extends BaseTest {
    protected static Chart chart;

    @BeforeAll
    static void getChart(){
        chart = presentation.getCharts().get(0);
    }

    @Test
    @DisplayName("获取柱状图列表")
    void testGetBarChartList(){
        List<BarChart> barCHartList = chart.getBarChartList();
        System.out.println(barCHartList.get(0));
    }

    @Test
    @DisplayName("获取指定柱状图")
    void testBarChartAt(){
        BarChart barChart = chart.barChartAt(0);
        System.out.println(barChart);
    }

    @Test
    @DisplayName("设置图表标题")
    void testSetTitle(){
        chart.setTitle("测试测试");
    }

    @Test
    @DisplayName("删除图表标题")
    void testGetTitle(){
        chart.removeTitle();
    }

    @Test
    @DisplayName("更新柱状图图表数据集")
    void testUpdateBarChart(){
        chart.updateBarChart(new Workbook("test.xlsx"), 0);
    }

}
