package com.livedrof.poi.excel.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.LinkedList;
import java.util.List;

/**
 * @param <T>
 * @author congsl
 */
public class Exporter<T> {
    private Workbook workbook;
    private List<DataSheet> sheets;
    public Exporter() {
        this.sheets = new LinkedList<>();
        this.workbook = new XSSFWorkbook();
    }

    public static <T> Exporter<T> getInstance() {
        return new Exporter<T>();
    }

    public <T> DataSheet<T> sheet(String title) {
        DataSheet<T> s = new DataSheet<>(title,this.workbook.createSheet(title));
        this.sheets.add(s);
        return s;
    }



    public Workbook toWorkbook() {
        for (DataSheet s : this.sheets) {
            s.renderSheet();
        }
        return workbook;
    }



    /**
     * 计算值
     *
     * @param <T>
     */
    public static interface ComputeValue<T> {
        public Object compute(T t);
    }
}
