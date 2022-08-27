package com.ice.office.domain.pptx.chart;

import com.ice.office.domain.xlsx.Workbook;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.util.List;
import java.util.stream.Collectors;

@Slf4j
@Data
public class Chart {
    // 图表对象
    private XSLFChart chart;

    // 图表对应的XML结点 <c:chart>
    private CTChart ctChart;

    // 图表所在的幻灯片对象
    private POIXMLDocumentPart parent;

    public Chart(XSLFChart chart) {
        this.chart = chart;
        this.ctChart = chart.getCTChart();
        this.parent = chart.getParent();
    }

    /**
     * 是否设置了图表标题
     *
     * @return 存在结果
     */
    public boolean isSetTitle() {
        return ctChart.isSetTitle();
    }

    /**
     * 设置图表标题
     *
     * @param title 图标标题
     */
    public void setTitle(String title) {
        this.chart.setTitleText(title);
    }

    public void removeTitle() {
        this.chart.removeTitle();
    }

    /**
     * 获取图表xml内所有柱状图xml标签
     *
     * @return 柱状图xml标签列表
     */
    public List<BarChart> getBarChartList() {
        return this.ctChart.getPlotArea().getBarChartList().stream().map(BarChart::new).collect(Collectors.toList());
    }

    /**
     * 获取图表xml内指定柱状图xml列表
     *
     * @return 柱状图xml标签
     */
    public BarChart barChartAt(int index) {
        return getBarChartList().get(index);
    }


    public void updateBarChart(Workbook workbook, int sheetIndex) {
        // TODO 更新柱状图
        chart.setWorkbook(workbook.getWorkbook());
        barChartAt(0).update(workbook.getSheetAt(sheetIndex));
    }

}
