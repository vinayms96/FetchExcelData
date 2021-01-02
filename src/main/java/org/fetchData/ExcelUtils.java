package org.fetchData;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

/**
 * Fetching the data from excel workbook
 */
public class ExcelUtils implements Constants {
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static Row row;
    private static Cell cell;
    private static ArrayList<String> rowListData = new ArrayList<>();
    private static HashMap<String, String> rowMapData = new HashMap<>();

    /**
     * Setting up the workbook through the path
     */
    public static void setupWorkbook() {
        try {
            workbook = new XSSFWorkbook(EXCELPATH);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Fetch the data from excel by passing sheetName &
     * adding the particular row data in ArrayList
     * @param sheetName - Sheet name to pick
     * @param dataName - Row name to pick
     * @return ArrayList<String>
     */
    public static ArrayList<String> getDataByList(String sheetName, String dataName) {
        sheet = workbook.getSheet(sheetName);
        Iterator<Row> rowIterator = sheet.iterator();
        row = rowIterator.next();

        Iterator<Cell> cellIterator = row.iterator();
        while (cellIterator.hasNext()) {
            cell = cellIterator.next();
            if (cell.getStringCellValue().equalsIgnoreCase("test_case")) {
                while (rowIterator.hasNext()) {
                    row = rowIterator.next();
                    if (row.getCell(cell.getColumnIndex()).getStringCellValue().equalsIgnoreCase(dataName)) {
                        while (cellIterator.hasNext()) {
                            cell = cellIterator.next();
                            if (row.getCell(cell.getColumnIndex()).getCellType().equals(CellType.STRING)) {
                                rowListData.add(row.getCell(cell.getColumnIndex()).getStringCellValue());
                            } else {
                                rowListData.add(NumberToTextConverter.toText(row.getCell(cell
                                        .getColumnIndex()).getNumericCellValue()));
                            }
                        }
                    }
                }
                return rowListData;
            }
        }
        return null;
    }

    /**
     * Fetch the data from Excel and add it to HashMap with the Column Name
     * and return Map data
     * @param sheetName
     * @param rowNum
     * @return HashMap<String, String>
     */
    public static HashMap<String, String> getDataByMap(String sheetName, int rowNum) {
        sheet = workbook.getSheet(sheetName);

        int cell_count = sheet.getRow(0).getLastCellNum();
        for(int cells = 0; cells < cell_count; cells++) {
            XSSFCell cell_data = sheet.getRow(rowNum).getCell(cells);
            if(cell_data.getCellType().equals(CellType.STRING)) {
                rowMapData.put(sheet.getRow(0).getCell(cells).getStringCellValue(),
                        cell_data.getStringCellValue());
            } else {
                rowMapData.put(sheet.getRow(0).getCell(cells).getStringCellValue(),
                        NumberToTextConverter.toText(cell_data.getNumericCellValue()));
            }
        }
        return rowMapData;
    }

    /**
     * Get the row count of specified sheet
     * @param sheetName
     * @return int
     */
    public static int getRowCount(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        return sheet.getLastRowNum();
    }

    /**
     * Get the cell count of specified sheet
     * @param sheetName
     * @return int
     */
    public static int getCellCount(String sheetName) {
        sheet = workbook.getSheet(sheetName);
        return sheet.getRow(0).getLastCellNum();
    }

    /**
     * Return the row data fetched from excel as List Data
     * @return ArrayList<String>
     */
    public static ArrayList<String> getRowListData() {
        return rowListData;
    }

    /**
     * Return the row data fetched from excel as Map Data
     * @return HashMap<String, String>
     */
    public static HashMap<String, String> getRowMapData() {
        return rowMapData;
    }
}
