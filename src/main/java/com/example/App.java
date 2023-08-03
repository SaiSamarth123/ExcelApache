package com.example;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    private String cs1Path;
    private String cs2Path;
    private String cs3Path;

    public App(String cs1Path, String cs2Path, String cs3Path) {
        this.cs1Path = cs1Path;
        this.cs2Path = cs2Path;
        this.cs3Path = cs3Path;
    }

    public void mergeData() throws IOException {
        Workbook cs1Workbook = new XSSFWorkbook(new FileInputStream(cs1Path));
        List<CS1Data> cs1DataList = readCS1Data(cs1Workbook);
        List<CS2Data> cs2DataList = readCS2Data(new XSSFWorkbook(new FileInputStream(cs2Path)));
        List<CS3Data> cs3DataList = readCS3Data(new XSSFWorkbook(new FileInputStream(cs3Path)));

        // Create map for CS2 and CS3 data for easier searching
        Map<String, CS2Data> cs2DataMap = new HashMap<>();
        for (CS2Data data : cs2DataList) {
            cs2DataMap.put(data.getVmName(), data);
        }

        Map<Integer, CS3Data> cs3DataMap = new HashMap<>();
        for (CS3Data data : cs3DataList) {
            cs3DataMap.put(data.getTechnicalAssetId(), data);
        }

        // Add new columns to CS1
        Sheet cs1Sheet = cs1Workbook.getSheetAt(0);
        Row row = cs1Sheet.getRow(0);
        int lastCellNum = row.getLastCellNum();
        row.createCell(lastCellNum).setCellValue("Owner Name");
        row.createCell(lastCellNum + 1).setCellValue("Service Id");
        row.createCell(lastCellNum + 2).setCellValue("Program Id");
        row.createCell(lastCellNum + 3).setCellValue("Program Name");

        // Fill new columns in CS1
        for (CS1Data cs1Data : cs1DataList) {
            CS2Data cs2Data = cs2DataMap.get(cs1Data.getServerName());
            if (cs2Data != null) {
                CS3Data cs3Data = cs3DataMap.get(cs2Data.getTechnicalAssetId());
                if (cs3Data != null) {
                    int rowIndex = cs1Data.getRowIndex();
                    row = cs1Sheet.getRow(rowIndex);
                    row.createCell(lastCellNum).setCellValue(cs3Data.getOwnerName());
                    row.createCell(lastCellNum + 1).setCellValue(cs3Data.getServiceId());
                    row.createCell(lastCellNum + 2).setCellValue(cs3Data.getProgramId());
                    row.createCell(lastCellNum + 3).setCellValue(cs3Data.getProgramName());
                }
            }
        }

        // Write changes to CS1
        try (FileOutputStream fileOut = new FileOutputStream(cs1Path)) {
            cs1Workbook.write(fileOut);
        }

        cs1Workbook.close();
    }

    private List<CS1Data> readCS1Data(Workbook workbook) {
        List<CS1Data> dataList = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String serverName = row.getCell(0).getStringCellValue();
            String labels = row.getCell(1).getStringCellValue();
            dataList.add(new CS1Data(serverName, labels, i));
        }
        return dataList;
    }

    private List<CS2Data> readCS2Data(Workbook workbook) {
        List<CS2Data> dataList = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String vmName = getStringValue(row.getCell(0));
            String vmType = getStringValue(row.getCell(1));
            Integer technicalAssetId = getIntegerValue(row.getCell(2));
            Integer productId = getIntegerValue(row.getCell(3));
            dataList.add(new CS2Data(vmName, vmType, technicalAssetId, productId));
        }
        return dataList;
    }

    private List<CS3Data> readCS3Data(Workbook workbook) {
        List<CS3Data> dataList = new ArrayList<>();
        Sheet sheet = workbook.getSheetAt(4);
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            String assetName = getStringValue(row.getCell(0));
            Integer technicalAssetId = getIntegerValue(row.getCell(1));
            String ownerName = getStringValue(row.getCell(2));
            Integer serviceId = getIntegerValue(row.getCell(3));
            Integer programId = getIntegerValue(row.getCell(4));
            String programName = getStringValue(row.getCell(5));
            dataList.add(new CS3Data(assetName, technicalAssetId, ownerName, serviceId, programId, programName));
        }
        return dataList;
    }

    private String getStringValue(Cell cell) {
        return cell == null || cell.getCellType() != CellType.STRING ? null : cell.getStringCellValue();
    }

    private Integer getIntegerValue(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) {
            return null;
        }
        double numericCellValue = cell.getNumericCellValue();
        return (numericCellValue == (int) numericCellValue) ? (int) numericCellValue : null;
    }

    public static void main(String[] args) {
        try {
            new App("CS1.xlsx", "CS2.xlsx", "CS3.xlsx").mergeData();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

class CS1Data {
    private String serverName;
    private String labels;
    private int rowIndex;

    public CS1Data(String serverName, String labels, int rowIndex) {
        this.serverName = serverName;
        this.labels = labels;
        this.rowIndex = rowIndex;
    }

    public String getServerName() {
        return serverName;
    }

    public void setServerName(String serverName) {
        this.serverName = serverName;
    }

    public String getLabels() {
        return labels;
    }

    public void setLabels(String labels) {
        this.labels = labels;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }
}

class CS2Data {
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

    public void setVmName(String vmName) {
        this.vmName = vmName;
    }

    public String getVmType() {
        return vmType;
    }

    public void setVmType(String vmType) {
        this.vmType = vmType;
    }

    public Integer getTechnicalAssetId() {
        return technicalAssetId;
    }

    public void setTechnicalAssetId(Integer technicalAssetId) {
        this.technicalAssetId = technicalAssetId;
    }

    public Integer getProductId() {
        return productId;
    }

    public void setProductId(Integer productId) {
        this.productId = productId;
    }
}

class CS3Data {
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

    public String getAssetName() {
        return assetName;
    }

    public void setAssetName(String assetName) {
        this.assetName = assetName;
    }

    public Integer getTechnicalAssetId() {
        return technicalAssetId;
    }

    public void setTechnicalAssetId(Integer technicalAssetId) {
        this.technicalAssetId = technicalAssetId;
    }

    public String getOwnerName() {
        return ownerName;
    }

    public void setOwnerName(String ownerName) {
        this.ownerName = ownerName;
    }

    public Integer getServiceId() {
        return serviceId;
    }

    public void setServiceId(Integer serviceId) {
        this.serviceId = serviceId;
    }

    public Integer getProgramId() {
        return programId;
    }

    public void setProgramId(Integer programId) {
        this.programId = programId;
    }

    public String getProgramName() {
        return programName;
    }

    public void setProgramName(String programName) {
        this.programName = programName;
    }
}
