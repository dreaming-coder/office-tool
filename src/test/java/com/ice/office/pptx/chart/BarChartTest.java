package com.ice.office.pptx.chart;

import com.ice.office.domain.pptx.chart.BarChart;
import com.ice.office.domain.xlsx.Sheet;
import com.ice.office.domain.xlsx.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.List;

public class BarChartTest extends ChartTest {
    protected static BarChart barChart;

    @BeforeAll
    static void getBarChart() {
        barChart = chart.getBarChartList().get(0);
    }

    @Test
    void testUpdate() {
        Workbook workbook = new Workbook("test.xlsx");
        Sheet sheet = workbook.getSheetAt(0);
        barChart.update(sheet);
    }

}
