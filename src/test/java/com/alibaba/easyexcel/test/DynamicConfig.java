package com.alibaba.easyexcel.test;

import java.util.Map;

/**
 * @Author: pengjianzhou
 * @Description:
 * @Date: Created in 下午4:36 18-12-28
 */
public class DynamicConfig {

    private Integer order;
    private String showHead;
    private String colName;

    private Boolean hasMapping = Boolean.FALSE;
    private Map<String, String> resultMapping;

    private Boolean hasDateFormat = Boolean.FALSE;
    private String formatStr;


    public DynamicConfig() {

    }

    public DynamicConfig(Integer order, String showHead, String colName) {
        this.order = order;
        this.showHead = showHead;
        this.colName = colName;
    }

    public Integer getOrder() {
        return order;
    }

    public void setOrder(Integer order) {
        this.order = order;
    }

    public String getShowHead() {
        return showHead;
    }

    public void setShowHead(String showHead) {
        this.showHead = showHead;
    }

    public String getColName() {
        return colName;
    }

    public void setColName(String colName) {
        this.colName = colName;
    }

    public Boolean getHasMapping() {
        return hasMapping;
    }

    public void setHasMapping(Boolean hasMapping) {
        this.hasMapping = hasMapping;
    }

    public Map<String, String> getResultMapping() {
        return resultMapping;
    }

    public void setResultMapping(Map<String, String> resultMapping) {
        this.resultMapping = resultMapping;
    }

    public Boolean getHasDateFormat() {
        return hasDateFormat;
    }

    public void setHasDateFormat(Boolean hasDateFormat) {
        this.hasDateFormat = hasDateFormat;
    }

    public String getFormatStr() {
        return formatStr;
    }

    public void setFormatStr(String formatStr) {
        this.formatStr = formatStr;
    }
}
