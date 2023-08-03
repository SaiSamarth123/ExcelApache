package com.example;

public class CS3Data {
    private String assetName;
    private Integer technicalAssetId;
    private String ownerName;
    private Integer serviceId;
    private Integer programId;
    private String programName;

    public CS3Data(String assetName, Integer technicalAssetId, String ownerName, Integer serviceId, Integer programId,
            String programName) {
        this.assetName = assetName;
        this.technicalAssetId = technicalAssetId;
        this.ownerName = ownerName;
        this.serviceId = serviceId;
        this.programId = programId;
        this.programName = programName;
    }

    public Integer getTechnicalAssetId() {
        return technicalAssetId;
    }

    public String getOwnerName() {
        return ownerName;
    }

    public Integer getServiceId() {
        return serviceId;
    }

    public Integer getProgramId() {
        return programId;
    }

    public String getProgramName() {
        return programName;
    }
}