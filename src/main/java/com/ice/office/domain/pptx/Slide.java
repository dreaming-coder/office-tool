package com.ice.office.domain.pptx;

import com.ice.office.domain.pptx.chart.Chart;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import java.util.LinkedList;
import java.util.List;

@Slf4j
@Data
public class Slide {
    private XSLFSlide slide;

    private List<Chart> charts;

    public Slide(XSLFSlide slide) {
        this.slide = slide;
        this.charts = getAllChartsFromSlide();
    }

    public List<POIXMLDocumentPart> getRelations() {
        return this.slide.getRelations();
    }

    /**
     * 获取指定幻灯片中的所有图表
     *
     * @return 该幻灯片中的图表List
     */
    private List<Chart> getAllChartsFromSlide() {
        List<Chart> charts = new LinkedList<>();
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                charts.add(new Chart((XSLFChart) relation));
            }
        }
        return charts;
    }

    /**
     * 获取指定幻灯片中的指定图表
     *
     * @return 该幻灯片中的图表List
     */
    public Chart getChartByIndex(int index) {
        return this.charts.get(index);
    }
}
