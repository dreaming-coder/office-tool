package com.ice.office.domain.pptx;

import com.ice.office.domain.pptx.chart.Chart;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;

@Slf4j
public final class Presentation {
    // 演示文稿
    private XMLSlideShow pptx;
    // 演示文稿中所有幻灯片
    private List<XSLFSlide> slides;

    public Presentation() {
    }

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
     * 获取演示文稿中所有幻灯片
     *
     * @return 幻灯片列表
     */
    public List<Slide> getSlides() {
        return pptx.getSlides().stream().map(Slide::new).collect(Collectors.toList());
    }

    /**
     * 获取演示文稿中所有图表
     *
     * @return 图表列表
     */
    public List<Chart> getCharts() {
        return pptx.getCharts().stream().map(Chart::new).collect(Collectors.toList());
    }

    /**
     * 获取演示文稿中指定索引的图表
     *
     * @param index 图表索引
     * @return 图表
     */
    public Chart chartAt(int index) {
        return getCharts().get(index);
    }

    /**
     * 获取指定索引的图表的xml标签
     *
     * @param index 图表索引
     * @return 图表的xml标签
     */
    public CTChart ctChartAt(int index) {
        return chartAt(index).getCtChart();
    }
}
