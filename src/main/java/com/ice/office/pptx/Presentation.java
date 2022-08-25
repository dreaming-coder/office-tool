package com.ice.office.pptx;

import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.LinkedList;
import java.util.List;

@Slf4j
@Data
public final class Presentation {
    // 演示文稿
    private final XMLSlideShow pptx;
    // 演示文稿中所有幻灯片
    private List<XSLFSlide> slides;

    /**
     * @param path 读取文件的路径
     */
    public Presentation(String path) {
        try {
            this.pptx = new XMLSlideShow(Files.newInputStream(Paths.get(path)));
            this.slides = pptx.getSlides();
            log.info("打开文件{}成功", path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * @param dest 输出文件的路径
     */
    public void writeTo(String dest) {
        try {
            this.pptx.write(Files.newOutputStream(new File(dest).toPath()));
            log.info("成功将文件写入{}", dest);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 获取指定幻灯片中的所有图表
     *
     * @param slide 指定幻灯片
     * @return 该幻灯片中的图表List
     */
    private List<XSLFChart> getAllChartsFromSlide(XSLFSlide slide) {
        List<XSLFChart> charts = new LinkedList<>();
        for (POIXMLDocumentPart relation : slide.getRelations()) {
            if (relation instanceof XSLFChart) {
                charts.add((XSLFChart) relation);
            }
        }
        return charts;
    }

    /**
     * 获取演示文稿中所有图表
     *
     * @return 图表列表
     */
    public List<XSLFChart> getCharts() {
        return pptx.getCharts();
    }

    /**
     * 获取指定索引的图表
     *
     * @param index 图表索引
     * @return 图表
     */
    public XSLFChart ChartAt(int index) {
        return getCharts().get(index);
    }

    /**
     * 获取指定索引的图表的xml标签
     *
     * @param index 图表索引
     * @return 图表的xml标签
     */
    public CTChart CTChartAt(int index) {
        return getCharts().get(index).getCTChart();
    }

    public void updateBarChart(CTBarChart barChart) {

    }

    void f() {
        XSLFChart chart = getCharts().get(0);
    }
}
