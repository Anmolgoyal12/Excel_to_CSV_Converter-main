package org.csdconverter;


import java.util.List;

/**
 * Class to represent sheet configuration details.
 */
public class SheetConfig {
    private final String sheetName;
    private final String csvName;
    private final Boolean isTranspose;
    private final Boolean isCommentRead;
    private final List<String> excludeFromTranspose;
    private final String outputDirectory;
    private final String range; // For Specific range of data

    public SheetConfig(String sheetName, String csvName, Boolean isTranspose, Boolean isCommentRead,String range,List<String> excludeFromTranspose, String outputDirectory) {
        this.sheetName = sheetName;
        this.csvName = csvName;
        this.isTranspose = isTranspose;
        this.isCommentRead = isCommentRead;
        this.excludeFromTranspose = excludeFromTranspose;
        this.outputDirectory = outputDirectory;
        this.range=range;
    }


    public String getSheetName() {
        return sheetName;
    }

    public String getCsvName() {
        return csvName;
    }

    public Boolean isTranspose() {
        return isTranspose;
    }

    public Boolean isCommentRead() {
        return isCommentRead;
    }

    public List<String> getExcludeFromTranspose() {
        return excludeFromTranspose;
    }

    public String getOutputDirectory() {
        return outputDirectory;
    }
    public  String getRange(){
        return  range;
    }

}
