package com.yuanjun;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.*;

import java.util.Date;


@Getter
@Setter
@EqualsAndHashCode
@ToString
public class DemoData {
    private String id;
    private String productName;
    private String caiZhi;
    private String guiGe;
    private String danWei;
    private String number;
    private String onePrice;
    private String Price;
    private String contractNumber;
    private String buyNumber;
    private String danNumber;
    private String date;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;

    public DemoData() {
    }

    public DemoData(String productName, String caiZhi, String guiGe, String danWei, String number, String onePrice, String price, String contractNumber) {
        this.productName = productName;
        this.caiZhi = caiZhi;
        this.guiGe = guiGe;
        this.danWei = danWei;
        this.number = number;
        this.onePrice = onePrice;
        Price = price;
        this.contractNumber = contractNumber;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getProductName() {
        return productName;
    }

    public void setProductName(String productName) {
        this.productName = productName;
    }

    public String getCaiZhi() {
        return caiZhi;
    }

    public void setCaiZhi(String caiZhi) {
        this.caiZhi = caiZhi;
    }

    public String getGuiGe() {
        return guiGe;
    }

    public void setGuiGe(String guiGe) {
        this.guiGe = guiGe;
    }

    public String getDanWei() {
        return danWei;
    }

    public void setDanWei(String danWei) {
        this.danWei = danWei;
    }

    public String getNumber() {
        return number;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public String getOnePrice() {
        return onePrice;
    }

    public void setOnePrice(String onePrice) {
        this.onePrice = onePrice;
    }

    public String getPrice() {
        return Price;
    }

    public void setPrice(String price) {
        Price = price;
    }

    public String getContractNumber() {
        return contractNumber;
    }

    public void setContractNumber(String contractNumber) {
        this.contractNumber = contractNumber;
    }

    public String getBuyNumber() {
        return buyNumber;
    }

    public void setBuyNumber(String buyNumber) {
        this.buyNumber = buyNumber;
    }

    public String getDanNumber() {
        return danNumber;
    }

    public void setDanNumber(String danNumber) {
        this.danNumber = danNumber;
    }

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }
}
