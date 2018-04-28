package com.livedrof.poi.excel.excel;

import java.util.LinkedList;
import java.util.List;

/**
 * 数据表格的元数据，由列(Column)构成
 */
public class DataTable {
    // 行数偏移量
    private Column root;

    private int rowOffset;

    private int colOffset;

    private int sumStartRowNum;

    int dataStartRowNum() {
        return this.rowOffset + this.getColsMaxDepth();
    }

    List<Column> computeColumns() {
        List<Column> resultColumns = new LinkedList<>();
        for (Column c : this.root.getChildren()) {
            this.computeCols(resultColumns, c);
        }
        return resultColumns;

    }

    private void computeCols(List<Column> result, Column col) {
        if (col.hasChildren()) {
            for (Column c : col.getChildren()) {
                computeCols(result, c);
            }
        } else {
            result.add(col);
        }
    }


    int getAbwRowOffset(Column column) {
        return this.rowOffset + column.getAbsRowOffset();
    }

    int getAbwColOffset(Column column) {
        return this.colOffset + column.getAbsColOffset();
    }

    /**
     * 所占用行数
     *
     * @param column
     * @return
     */
    int getRowsNum(Column column) {
        return column.getRowsNum(this.getColsMaxDepth()) - 1;
    }


    int getColsNum(Column column) {
        return column.getColsNum(column) - 1;
    }

    DataTable() {
        this.root = new Column();
        this.colOffset = 0;
    }

    public Column addColumn(String title, String fieldCode) {
        this.root.nextColumn(title, fieldCode);
        return this.root;
    }

    /**
     * 取值
     *
     * @param title
     * @param computeValueFunc
     * @return
     */
    public Column addColumn(String title, Exporter.ComputeValue computeValueFunc) {
        this.root.addColumn(title, computeValueFunc);
        return this.root;
    }

    /**
     * 返回Column的列表
     *
     * @return
     */
    List<Column> getCols() {
        return this.root.getChildren();
    }

    public int getColsMaxDepth() {
        return this.getColsMaxDepth(this.root, 0);
    }

    private int getColsMaxDepth(Column col, int depth) {
        int maxDepth = depth;
        if (col.hasChildren()) {
            for (Column childCol : col.getChildren()) {
                int childDepth = this.getColsMaxDepth(childCol, depth);
                if (childDepth > maxDepth) {
                    maxDepth = childDepth;
                }
            }
        } else {
            maxDepth = maxDepth > col.getDepth() ? maxDepth : col.getDepth();
        }
        return maxDepth;
    }

    public void setRowOffset(int rowOffset) {
        this.rowOffset = rowOffset;
    }

    public void setSumStartRowNum(int sumStartRowNum) {
        this.sumStartRowNum = sumStartRowNum;
    }

    public int getSumStartRowNum() {
        return sumStartRowNum;
    }
}
