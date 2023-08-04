package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.checkerframework.common.returnsreceiver.qual.This;

import com.google.common.collect.ArrayListMultimap;
import com.google.common.collect.ListMultimap;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

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

        ListMultimap<String, CS2Data> cs2DataMap = ArrayListMultimap.create();
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
            List<CS2Data> matchingCS2Data = cs2DataMap.get(cs1Data.getServerName());
            for (CS2Data cs2Data : matchingCS2Data) {
                CS3Data cs3Data = cs3DataMap.get(cs2Data.getTechnicalAssetId());
                if (cs3Data != null) {
                    int rowIndex = cs1Data.getRowIndex();
                    row = cs1Sheet.createRow(rowIndex);
                    row.createCell(0).setCellValue(cs1Data.getServerName());
                    row.createCell(1).setCellValue(cs1Data.getLabels());
                    row.createCell(lastCellNum).setCellValue(cs3Data.getOwnerName());
                    row.createCell(lastCellNum + 1).setCellValue(cs3Data.getServiceId().toString());
                    row.createCell(lastCellNum + 2).setCellValue(cs3Data.getProgramId().toString());
                    row.createCell(lastCellNum + 3).setCellValue(cs3Data.getProgramName());
                    System.out.println("Added item to CS1");
                }
            }
        }

        // Write changes to CS1
        try (FileOutputStream fileOut = new FileOutputStream(cs1Path)) {
            cs1Workbook.write(fileOut);
            System.out.println("Successfully");
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
        Sheet sheet = workbook.getSheetAt(3);
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
        // if (args.length != 3) {
        // System.err.println("Please provide paths to CS1, CS2, and CS3 files as
        // arguments.");
        // System.exit(1);
        // }
        // try {
        // new App(args[0], args[1], args[2]).mergeData();
        // } catch (IOException e) {
        // e.printStackTrace();
        // }
    }
}
