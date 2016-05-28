/*
 *  Copy right 2016 pamarin.com
 */
package com.pamarin.util.excel.integrator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author jittagornp
 */
public class ExcelSheetIntegrator {

    private List<ExcelFile> inputFiles;
    private File outputFile;
    private final Map<Integer, CellStyle> cellStyleMap;

    private ExcelSheetIntegrator() {
        cellStyleMap = new HashMap<>();
    }

    public static ExcelSheetIntegrator newInstance() {
        return new ExcelSheetIntegrator();
    }

    public ExcelSheetIntegrator addExcelFile(ExcelFile excelFile) {

        if (excelFile == null) {
            throw new NullPointerException("require excelFile");
        }

        getInputFiles().add(excelFile);
        return this;
    }

    public ExcelSheetIntegrator toTargetFile(File file) {

        if (file == null) {
            throw new NullPointerException("require file");
        }

        this.outputFile = file;
        return this;
    }

    private List<ExcelFile> getInputFiles() {

        if (inputFiles == null) {
            inputFiles = new ArrayList<>();
        }

        return inputFiles;
    }

    public File integrate() throws IOException {
        Workbook destWorkbook = createWorkbook(outputFile);
        try (OutputStream outputStream = new FileOutputStream(outputFile)) {
            List<ExcelFile> excelFiles = getInputFiles();
            for (ExcelFile excelFile : excelFiles) {
                integrateSheet(excelFile, destWorkbook);
            }

            destWorkbook.write(outputStream);
        }

        return this.outputFile;
    }

    private void integrateSheet(ExcelFile excelFile, Workbook destWorkbook) throws IOException {
        File file = excelFile.getFile();
        List<String> sheetNames = excelFile.getSheetNames();

        try (InputStream inputStream = new FileInputStream(file)) {
            Workbook srcWorkbook = loadWorkbook(file, inputStream);
            copyWorkbook(srcWorkbook, destWorkbook, sheetNames);
        }
    }

    private void copyWorkbook(Workbook srcWorkbook, Workbook destWorkbook, List<String> sheetNames) {
        int numberOfSheets = srcWorkbook.getNumberOfSheets();
        for (int index = 0; index < numberOfSheets; index++) {
            copySheets(srcWorkbook, destWorkbook, sheetNames, index);
        }
    }

    private boolean hasData(Sheet sheet) {
        return sheet.iterator().hasNext();
    }

    private void copySheets(Workbook srcWorkbook, Workbook destWorkbook, List<String> sheetNames, int index) {
        Sheet srcSheet = srcWorkbook.getSheetAt(index);
        if (!hasData(srcSheet)) {
            return;
        }

        String sheetName;
        try {
            sheetName = sheetNames.get(index);
        } catch (IndexOutOfBoundsException ex) {
            sheetName = srcSheet.getSheetName();
        }

        Sheet destSheet = createSheet(destWorkbook, sheetName);
        copySheet(srcSheet, destSheet);
        copyMergedRegion(srcSheet, destSheet);
    }

    private void copyMergedRegion(Sheet srcSheet, Sheet destSheet) {
        try {
            int numb = srcSheet.getNumMergedRegions();
            for (int index = 0; index < numb; index++) {
                destSheet.addMergedRegion(srcSheet.getMergedRegion(index));
            }
        } catch (Exception ex) {

        }
    }

    private boolean isAlreadyExistSheet(IllegalArgumentException ex) {
        return ex.getMessage()
                .equals("The workbook already contains a sheet of this name");
    }

    private Sheet createSheet(Workbook workbook, String sheetName) {
        Sheet sheet = null;
        int numb = 1;
        while (true) {
            try {
                sheet = workbook.createSheet(sheetName);
                break;
            } catch (IllegalArgumentException ex) {
                if (isAlreadyExistSheet(ex)) {
                    sheetName = sheetName + "-" + numb;
                    numb = numb + 1;
                } else {
                    break;
                }
            }
        }

        return sheet;
    }

    private void copySheet(Sheet srcSheet, Sheet destSheet) {
        Iterator<Row> iterator = srcSheet.iterator();
        while (iterator.hasNext()) {
            Row srcRow = iterator.next();
            Row destRow = destSheet.createRow(srcRow.getRowNum());
            copyRow(srcRow, destRow);
        }
    }

    private void copyRow(Row srcRow, Row destRow) {
        Iterator<Cell> iterator = srcRow.iterator();
        while (iterator.hasNext()) {
            Cell srcCell = iterator.next();
            Cell destCell = destRow.createCell(srcCell.getColumnIndex());
            copyCell(srcCell, destCell);
        }

        destRow.setHeight(srcRow.getHeight());
        copyRowStyle(srcRow, destRow);
    }

    private void copyCell(Cell srcCell, Cell destCell) {
        copyValue(srcCell, destCell);
        copyCellStyle(srcCell, destCell);
        copyCellWidth(srcCell, destCell);
    }

    private void copyCellWidth(Cell srcCell, Cell destCell) {
        int columnWidth = srcCell.getRow().getSheet().getColumnWidth(srcCell.getColumnIndex());
        destCell.getRow().getSheet().setColumnWidth(destCell.getColumnIndex(), columnWidth);
    }

    private void copyValue(Cell srcCell, Cell destCell) {
        switch (srcCell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                destCell.setCellValue(srcCell.getRichStringCellValue());
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                destCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case HSSFCell.CELL_TYPE_ERROR:
                destCell.setCellErrorValue(srcCell.getErrorCellValue());
                break;
            case HSSFCell.CELL_TYPE_FORMULA:
                destCell.setCellFormula(srcCell.getCellFormula());
                break;
            default:
                break;
        }
    }

    private void copyCellStyle(Cell srcCell, Cell destCell) {

        if (srcCell.getCellStyle() == null) {
            return;
        }

        int hashCode = srcCell.getCellStyle().hashCode();
        CellStyle cellStyle = cellStyleMap.get(hashCode);
        if (cellStyle == null) {
            cellStyle = destCell.getSheet().getWorkbook().createCellStyle();
            cellStyle.cloneStyleFrom(srcCell.getCellStyle());
            cellStyleMap.put(hashCode, cellStyle);
        }

        destCell.setCellStyle(cellStyle);
    }

    private void copyRowStyle(Row srcRow, Row destRow) {

        if (srcRow.getRowStyle() == null) {
            return;
        }

        int hashCode = srcRow.getRowStyle().hashCode();
        CellStyle cellStyle = cellStyleMap.get(hashCode);
        if (cellStyle == null) {
            cellStyle = destRow.getSheet().getWorkbook().createCellStyle();
            cellStyle.cloneStyleFrom(srcRow.getRowStyle());
            cellStyleMap.put(hashCode, cellStyle);
        }

        destRow.setRowStyle(cellStyle);
    }

    private boolean isVersion2003(File file) {
        return file.getName().endsWith(".xls");
    }

    private Workbook loadWorkbook(File file, InputStream inputStream) throws IOException {
        if (isVersion2003(file)) {
            return new HSSFWorkbook(inputStream);
        } else { //2007+
            return new XSSFWorkbook(inputStream);
        }
    }

    private Workbook createWorkbook(File file) throws IOException {
        if (isVersion2003(file)) {
            return new HSSFWorkbook();
        } else { //2007+
            return new XSSFWorkbook();
        }
    }
}
