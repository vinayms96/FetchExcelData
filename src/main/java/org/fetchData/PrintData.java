package org.fetchData;

/**
 * Print the data fetched from excel
 */
public class PrintData
{
    /**
     * Just to print and check the data retrieved
     * @param args
     */
    public static void main( String[] args )
    {
        ExcelUtils.setupWorkbook();
        System.out.println( ExcelUtils.getDataByList("customerData","user_test1") );
        ExcelUtils.getDataArray().clear();
        System.out.println( ExcelUtils.getDataByList("customerData","user_test2") );
        System.out.println( "Row Count = " + ExcelUtils.getRowCount("customerData") );
        System.out.println( "Cell Count = " + ExcelUtils.getCellCount("customerData") );

        ExcelUtils.getDataArray().clear();
        System.out.println( ExcelUtils.getDataByList("productData","simple_product") );
        ExcelUtils.getDataArray().clear();
        System.out.println( ExcelUtils.getDataByList("productData","config_product") );
        System.out.println( "Row Count = " + ExcelUtils.getRowCount("productData") );
        System.out.println( "Cell Count = " + ExcelUtils.getCellCount("productData") );
    }
}
