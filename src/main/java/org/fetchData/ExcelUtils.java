package org.fetchData;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Iterator;

/**
 * Fetching the data from excel workbook
 */
public class ExcelUtils implements Constants {
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;
    private static Row row;
    private static Cell cell;
    private static ArrayList<String> rowData = new ArrayList<>();

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
     * Fetch the data from excel by passing sheetname &
     * adding the particular row data in ArrayList
     * @param sheetName
     * @param dataName
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
                                rowData.add(row.getCell(cell.getColumnIndex()).getStringCellValue());
                            } else {
                                rowData.add(NumberToTextConverter.toText(row.getCell(cell
                                        .getColumnIndex()).getNumericCellValue()));
                            }
                        }
                    }
                }
                return rowData;
            }
        }
        return null;
    }

    public static void getDataByMap(String sheetName) {
        
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
     * Return the row data fetched from excel
     * @return ArrayList<String>
     */
    public static ArrayList<String> getDataArray() {
        return rowData;
    }
}
