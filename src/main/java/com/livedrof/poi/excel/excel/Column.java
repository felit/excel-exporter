package com.livedrof.poi.excel.excel;

import java.util.LinkedList;
import java.util.List;

/**
 * 每一列的元数据信息
 */
public class Column {

    //列的展现与计算
    private String title;
    private String field;
    private Exporter.ComputeValue computeValFunc;
    //列的位置
    private Integer rowOffset;
    private Integer colOffset;
    private Integer depth;
    //
    private Column parent;
    //子列
    private List<Column> children;

    public Column getParent() {
        return parent;
    }

    public Column() {
        this.initOffsetAndDepth(null);
        this.parent = null;
    }

    private Column(Column parent, String title) {
        this.initOffsetAndDepth(parent);
        this.title = title;
    }

    private Column(Column parent, String title, String field) {
        this.initOffsetAndDepth(parent);
        this.title = title;
        this.field = field;
    }

    private Column(Column parent, String title, Exporter.ComputeValue computeValueFunc) {
        this.initOffsetAndDepth(parent);
        this.title = title;
        this.computeValFunc = computeValueFunc;
    }

    /**
     * 初始化坐标信息(rowOffset colOffset depth等)
     *
     * @param parent
     */
    private void initOffsetAndDepth(Column parent) {
        this.parent = parent;
        if (this.parent == null) {
            this.rowOffset = 0;
            this.colOffset = 0;
            this.depth = 0;
        } else {
            this.rowOffset = 1;
            //TODO 与添加顺序有关
            this.colOffset = this.getColsNum(parent);
            this.depth = parent.getDepth() + 1;
        }
    }


    public boolean hasField() {
        return this.field != null;
    }

    @Deprecated
    private Integer getRowOffset() {
        return rowOffset == null ? 0 : rowOffset;
    }

    @Deprecated
    private Integer getColOffset() {
        return colOffset == null ? 0 : colOffset;
    }

    public Integer getDepth() {
        return depth == null ? 0 : depth;
    }

    /**
     * 确定单元格横跨几列
     *
     * @return
     */
    public Integer getColsNum(Column col) {
        int cols = 0;
        if (col.hasChildren()) {
            for (Column childCol : col.getChildren()) {
                cols += childCol.getColsNum(childCol);
            }
        } else {
            cols = 1;
        }
        return cols;

    }

    public Integer getAbsRowOffset() {
        if (this.getParent().isRoot()) {
            return this.rowOffset;
        } else {
            return this.returnParent().getAbsRowOffset() + this.rowOffset;
        }
    }

    public Integer getAbsColOffset() {
        if (this.getParent().isRoot()) {
            return this.colOffset;
        } else {
            return this.returnParent().getAbsColOffset() + this.colOffset;
        }

    }

    /**
     * 合并行数
     *
     * @param treeDepth
     * @return
     */
    public Integer getRowsNum(Integer treeDepth) {
        if (this.hasChildren()) {
            return 1;
        } else {
            return treeDepth - depth + 1;
        }

    }


    public Column addChild(Column column) {
        this.getChildren().add(column);
        return this;
    }

    public Column nextColumn(String title, String fieldCode) {
        this.getChildren().add(new Column(this, title, fieldCode));
        return this;
    }

    public Column addColumn(String title, Exporter.ComputeValue computeValueFunc) {
        this.getChildren().add(new Column(this, title, computeValueFunc));
        return this;
    }


    public Column parentColumn(String title) {
        Column parentColumn = new Column(this, title);
        this.addChild(parentColumn);
        return parentColumn;
    }

    public Column returnParent() {
        return this.parent;
    }

    public boolean hasChildren() {
        return this.children != null;
    }

    private boolean isRoot() {
        return this.parent == null;
    }

    public List<Column> getChildren() {
        if (this.children == null) {
            this.children = new LinkedList<>();
        }
        return this.children;
    }

    @Override
    public String toString() {
        return "Column{" +
                "title='" + title + '\'' +
                ", field='" + field + '\'' +
                ", computeValFunc=" + computeValFunc +
                ", rowOffset=" + rowOffset +
                ", colOffset=" + colOffset +
                ", depth=" + depth +
                ", children=" + children +
                '}';
    }

    public String getTitle() {
        return title;
    }

    public String getField() {
        return field;
    }

    public Exporter.ComputeValue getComputeValFunc() {
        return computeValFunc;
    }
}
