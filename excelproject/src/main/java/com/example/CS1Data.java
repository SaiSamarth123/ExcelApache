package com.example;

public class CS1Data {
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

    public String getLabels() {
        return labels;
    }

    public int getRowIndex() {
        return rowIndex;
    }
}
