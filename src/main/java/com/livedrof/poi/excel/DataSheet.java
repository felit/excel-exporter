package com.livedrof.poi.excel;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

/**
 * 代表一下个Sheet页
 *
 * @param <T>
 * @author jacky
 */
public class DataSheet<T> {

    private DataTable dataTable;
    private List<Summary> headers;
    private List<T> data;
    private T sumData;
    private String title;
    private Sheet sheet;

    public DataSheet(String title, Sheet sheet) {
        this.headers = new LinkedList<>();
        this.title = title;
        this.sheet = sheet;
    }

    /**
     * 设置明细数据
     *
     * @param data
     * @return
     */
    public DataSheet<T> setData(List<T> data) {
        this.data = data;
        return this;
    }

    /**
     * 设置合计数据
     *
     * @param sumData
     * @return
     */
    public DataSheet<T> setSumData(T sumData) {
        this.sumData = sumData;
        return this;
    }


    /**
     * @param data 数据，其中最后一行为合计行
     * @return
     */
    public DataSheet<T> setDataWithSum(List<T> data) {
        this.sumData = data.remove(data.size() - 1);
        this.data = data;
        return this;
    }

    public DataSheet<T> addHeader(String title) {
        this.headers.add(new Summary(title));
        return this;
    }


    public DataTable getDataTable() {
        if (this.dataTable == null) {
            this.dataTable = new DataTable();
        }
        return this.dataTable;
    }


    public void renderSheet() {
        this.renderHeader();
        this.renderTableHeader();
        this.renderData();
        if (this.sumData != null) {
            this.renderSum();
        }
    }

    /**
     * 渲染页头
     */
    private void renderTableHeader() {
        CellStyle headerStyle = this.getTableHeaderStyle();
        DataTable table = this.getDataTable();
        //TODO 外部设置
        table.setRowOffset(5);

        for (Column col : table.getCols()) {
            this.innerRenderTableHeader(table, headerStyle, col);
        }
    }

    /**
     * 渲染页头
     */
    private void renderHeader() {
        CellStyle headerStyle = this.getHeaderStyle();
        for (int i = 0; i < this.headers.size(); i++) {
            Summary header = this.headers.get(i);
            Row headerRow = sheet.createRow(i);
            this.addCell(headerRow, 0, header.getTitle(), headerStyle);
            sheet.addMergedRegion(new CellRangeAddress(i, i, 0, 5));
        }
    }

    /**
     * @param table
     * @param headerStyle
     * @param col
     */
    private void innerRenderTableHeader(DataTable table, CellStyle headerStyle, Column col) {
        int rowOffset = table.getAbwRowOffset(col);
        int colOffset = table.getAbwColOffset(col);
        Row row = sheet.getRow(rowOffset);
        if (row == null) {
            row = sheet.createRow(rowOffset);
        }
        addCell(row, colOffset, col.getTitle(), headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowOffset, rowOffset + table.getRowsNum(col), colOffset, colOffset + table.getColsNum(col)));
        if (col.hasChildren()) {
            for (Column childCol : col.getChildren()) {
                this.innerRenderTableHeader(table, headerStyle, childCol);
            }
        }
    }

    /**
     * 渲染数据
     */
    private void renderData() {
        List<Column> columns = this.dataTable.computeColumns();
        int dataRowOffset = this.dataTable.dataStartRowNum();
        //TODO 如果this.data为null显示错误，如果为空列表则显示info级日志
        if (this.data != null && this.data.size() > 0) {
            for (T rowData : this.data) {
                dataRowOffset++;
                Row row = sheet.getRow(dataRowOffset);
                if (row == null) {
                    row = sheet.createRow(dataRowOffset);
                }
                for (Column c : columns) {
                    Object val = this.getVal(rowData, c);
                    addCell(row, c.getAbsColOffset(), val, this.getDataCellStyle(val));
                }
            }
        }

        this.dataTable.setSumStartRowNum(dataRowOffset);
    }

    /**
     * 计算对t对就某一列的值
     *
     * @param t
     * @param column
     * @return
     */
    private Object getVal(T t, Column column) {
        Object val = null;
        if (column.hasField()) {
            try {
//                log.debug("column:{}", column);
                val = beanGetValue(t, column.getField());
            } catch (Exception e) {
                e.printStackTrace();
            }
        } else {
            val = column.getComputeValFunc().compute(t);
        }
        return val;
    }

    /**
     * 渲染合计信息
     */
    private void renderSum() {
        int sumStartRowNum = this.dataTable.getSumStartRowNum() + 2;
        List<Column> columns = this.dataTable.computeColumns();
        Row row = sheet.getRow(sumStartRowNum);
        if (row == null) {
            row = sheet.createRow(sumStartRowNum);
        }
        for (Column c : columns) {
            Object val = this.getVal(this.sumData, c);
            if (c.getAbsColOffset() == 0) {
                addCell(row, c.getAbsColOffset(), "合计：", this.getDataCellStyle(""));
            } else {
                if (val == null) {
                    val = "--";
                }
                addCell(row, c.getAbsColOffset(), val, this.getDataCellStyle(val));
            }
        }
    }


    public static Object beanGetValue(Object object, String fieldName) throws Exception {
        Class<?> clazz = object.getClass();
        Field field = clazz.getDeclaredField(fieldName);
        PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
        Method getMethod = pd.getReadMethod();
        return getMethod.invoke(object);
    }

    /**
     * 设置表头样式
     *
     * @return
     */

    private CellStyle getTableHeaderStyle() {
        CellStyle titleStyle = this.getWorkbook().createCellStyle();
        titleStyle.setBorderRight(CellStyle.BORDER_THIN); // 设置边框线
        titleStyle.setRightBorderColor(IndexedColors.GREY_80_PERCENT.getIndex());
        titleStyle.setBorderLeft(CellStyle.BORDER_THIN);
        titleStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        titleStyle.setBorderTop(CellStyle.BORDER_THIN);
        titleStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        titleStyle.setBorderBottom(CellStyle.BORDER_THIN);
        titleStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        titleStyle.setAlignment(CellStyle.ALIGN_CENTER);// 设置居中
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 设置居中
        titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        Font titleFont = this.getWorkbook().createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 10);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        titleStyle.setFont(titleFont);
        return titleStyle;
    }

    public Workbook getWorkbook() {
        return this.sheet.getWorkbook();
    }

    /**
     * 页头格式
     *
     * @return
     */
    private CellStyle getHeaderStyle() {
        CellStyle titleStyle = this.getWorkbook().createCellStyle();
        titleStyle.setVerticalAlignment(CellStyle.ALIGN_LEFT);
        Font titleFont = this.getWorkbook().createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 10);
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        titleStyle.setFont(titleFont);
        return titleStyle;
    }

    private CellStyle getDataCellStyle(Object val) {
        CellStyle cellStyle = this.getWorkbook().createCellStyle();
        // 设置边框线
        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
        cellStyle.setRightBorderColor(IndexedColors.GREY_80_PERCENT.getIndex());
        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        cellStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
        cellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        cellStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);// 设置居中
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);// 设置居中
        Font titleFont = this.getWorkbook().createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 10);
        cellStyle.setFont(titleFont);

        if (val instanceof Float || val instanceof Double || val instanceof BigDecimal) {
            DataFormat df = this.getWorkbook().createDataFormat();
            cellStyle.setDataFormat(df.getFormat("#,#0.00"));
        }
        return cellStyle;
    }

    /**
     * Copy原有代码
     *
     * @param row
     * @param column
     * @param val
     * @param style
     * @return
     */
    private static Cell addCell(Row row, int column, Object val, CellStyle style) {
        Cell cell = row.createCell(column);
        try {
            if (val == null) {
                cell.setCellValue("");
            } else if (val instanceof String) {
                cell.setCellValue((String) val);
            } else if (val instanceof Integer) {
                cell.setCellValue((Integer) val);
            } else if (val instanceof Long) {
                cell.setCellValue((Long) val);
            } else if (val instanceof Double) {
//                cell.setCellValue(MathUtils.doubleFormat((Double) val, false));
//            } else if (val instanceof Float) {
//                cell.setCellValue((Float) val);
//            } else if (val instanceof Date) {
//                cell.setCellValue(DateUtils.formatDate((Date) val, DateUtils.FORMAT3));
//            } else if (val instanceof LocalDate) {
//                cell.setCellValue(((LocalDate) val).toString());
//            } else if (val instanceof LocalTime) {
//                cell.setCellValue(((LocalTime) val).toString());
//            } else if (val instanceof BigDecimal) {
//                cell.setCellValue(MathUtils.doubleFormat(((BigDecimal) val).doubleValue(), false));
            } else if (val instanceof UUID) {
                cell.setCellValue((val).toString());
            }
        } catch (Exception ex) {
//TODO 需要处理
//            cell.setCellValue(JSONObject.toJSONString(val));
        }

        cell.setCellStyle(style);
        return cell;
    }

}
