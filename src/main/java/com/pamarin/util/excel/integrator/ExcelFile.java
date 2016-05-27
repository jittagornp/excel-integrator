/*
 *  Copy right 2016 pamarin.com
 */
package com.pamarin.util.excel.integrator;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * @author jittagornp
 */
public class ExcelFile {

    private final File file;

    private List<String> sheetNames;

    private ExcelFile(File file) {
        this.file = file;
    }

    public static ExcelFile fromFile(File file) {
        return new ExcelFile(file);
    }

    public ExcelFile withSheetName(String sheetName) {
        getSheetNames().add(sheetName);
        return this;
    }

    public File getFile() {
        return file;
    }

    public List<String> getSheetNames() {

        if (sheetNames == null) {
            sheetNames = new ArrayList<>();
        }

        return sheetNames;
    }
}
