package com.ice.office.domain.pptx.chart;

import com.ice.office.domain.xlsx.Sheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.List;
import java.util.Map;

public class BarChart {
    private CTBarChart ctBarChart;

    public BarChart(CTBarChart ctBarChart) {
        this.ctBarChart = ctBarChart;
    }

    /**
     * 更新图表数据
     *
     * @param sheet 需要更新的表格sheet对象
     */
    public void update(Sheet sheet) {
        // sheet名
        String sheetName = sheet.getSheetName();

        List<Map<String, String>> series = sheet.getSeriesArray();
        List<Map<String, String>> categories = sheet.getCategoriesArray();
        List<List<Map<String, String>>> serValList = sheet.getSerValList();

        if (serValList == null || serValList.size() == 0) return;
        int seriesNum = series.size();
        int categoriesNum = categories.size();
        List<CTBarSer> serList = ctBarChart.getSerList();

        if (ctBarChart.isSetExtLst()) {
            ctBarChart.unsetExtLst();
        }

        int existSerNum = serList.size();
        int i = 0;
        for (; i < seriesNum; i++) {
            if (i < existSerNum) {
                CTBarSer ctBarSer = serList.get(i);


                ctBarSer.getTx().getStrRef().setF(sheetName + "!" + series.get(i).get("address"));
                ctBarSer.getTx().getStrRef().getStrCache().getPtList().get(0).setV(series.get(i).get("value"));

                //  更新categories
                CTStrRef ctStrRef = ctBarSer.getCat().getStrRef();
                updateCat(sheetName, categories, categoriesNum, ctStrRef);

                // 更新value
                CTNumRef ctNumRef = ctBarSer.getVal().getNumRef();
                updateVal(sheetName, serValList, categoriesNum, i, ctNumRef);

            } else {
                CTBarSer ctBarSer = ctBarChart.addNewSer();
                ctBarSer.addNewIdx().setVal(i);
                ctBarSer.addNewOrder().setVal(i);
                CTSerTx tx = ctBarSer.addNewTx();
                CTStrRef strRef = tx.addNewStrRef();
                strRef.setF(sheetName + "!" + series.get(i).get("address"));
                CTStrData ctStrData = strRef.addNewStrCache();
                ctStrData.addNewPtCount().setVal(1);
                CTStrVal pt = ctStrData.addNewPt();
                pt.setIdx(0);
                pt.setV(series.get(i).get("value"));

                CTAxDataSource cat = ctBarSer.addNewCat();
                CTStrRef ctStrRef = cat.addNewStrRef();
                updateCat(sheetName, categories, categoriesNum, ctStrRef);

                CTNumDataSource val = ctBarSer.addNewVal();
                CTNumRef ctNumRef = val.addNewNumRef();
                updateVal(sheetName, serValList, categoriesNum, i, ctNumRef);
            }
        }

        // 删除多余的series缓存
        if (i < existSerNum){
            for (int j = i; j < existSerNum; j++) {
                ctBarChart.removeSer(j);
            }
        }
    }

    private void updateVal(String sheetName, List<List<Map<String, String>>> serValList, int categoriesNum, int i, CTNumRef ctNumRef) {
        ctNumRef.setF(sheetName + "!" + serValList.get(i).get(0).get("address") + ":" + serValList.get(i).get(categoriesNum - 1).get("address"));
        if (ctNumRef.isSetNumCache()) {
            ctNumRef.unsetNumCache();
        }
        CTNumData ctNumData = ctNumRef.addNewNumCache();
        ctNumData.setFormatCode("General");
        ctNumData.addNewPtCount().setVal(categoriesNum);
        for (int k = 0; k < categoriesNum; k++) {
            Map<String, String> serVal = serValList.get(i).get(k);
            CTNumVal pt = ctNumData.addNewPt();
            pt.setIdx(k);
            pt.setV(serVal.get("value"));
        }
    }

    private void updateCat(String sheetName, List<Map<String, String>> categories, int categoriesNum, CTStrRef ctStrRef) {
        ctStrRef.setF(sheetName + "!" + categories.get(0).get("address") + ":" + categories.get(categoriesNum - 1).get("address"));
        if (ctStrRef.isSetStrCache()) {
            ctStrRef.unsetStrCache();
        }

        CTStrData ctStrData = ctStrRef.addNewStrCache();
        ctStrData.addNewPtCount().setVal(categoriesNum);
        for (int i = 0; i < categoriesNum; i++) {
            CTStrVal pt = ctStrData.addNewPt();
            Map<String, String> category = categories.get(i);
            pt.setIdx(i);
            pt.setV(category.get("value"));
        }
    }

    @Override
    public String toString() {
        return "BarChart{" +
                "ctBarChart=" + ctBarChart +
                '}';
    }
}
