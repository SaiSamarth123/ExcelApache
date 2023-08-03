package com.example;

public class CS2Data {
    private String vmName;
    private String vmType;
    private Integer technicalAssetId;
    private Integer productId;

    public CS2Data(String vmName, String vmType, Integer technicalAssetId, Integer productId) {
        this.vmName = vmName;
        this.vmType = vmType;
        this.technicalAssetId = technicalAssetId;
        this.productId = productId;
    }

    public String getVmName() {
        return vmName;
    }

    public Integer getTechnicalAssetId() {
        return technicalAssetId;
    }
}
