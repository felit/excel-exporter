package com.livedrof.poi.excel;

/**
 * 摘要信息,目前只支持文本，合并默认为5个。
 */
class Summary {
    private String title;

    public Summary(String title) {
        this.title = title;
    }

    public String getTitle() {
        return title;
    }
}
