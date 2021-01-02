package org.fetchData;

/**
 * Print the data fetched from excel
 */
public class PrintData
{
    /**
     * Print and check the data retrieved
     * @param args
     */
    public static void main( String[] args )
    {
        ExcelUtils.setupWorkbook();
        System.out.println( ExcelUtils.getDataByList("customerData","user_test1") );
        ExcelUtils.getRowListData().clear();
        System.out.println( ExcelUtils.getDataByList("customerData","user_test2") );
        System.out.println( "Row Count = " + ExcelUtils.getRowCount("customerData") );
        System.out.println( "Cell Count = " + ExcelUtils.getCellCount("customerData") );

        System.out.println();

        ExcelUtils.getRowListData().clear();
        System.out.println( ExcelUtils.getDataByList("productData","simple_product") );
        ExcelUtils.getRowListData().clear();
        System.out.println( ExcelUtils.getDataByList("productData","config_product") );
        System.out.println( "Row Count = " + ExcelUtils.getRowCount("productData") );
        System.out.println( "Cell Count = " + ExcelUtils.getCellCount("productData") );

        System.out.println();

        System.out.println(ExcelUtils.getDataByMap("customerData", 1));
        ExcelUtils.getRowMapData().clear();
        System.out.println( ExcelUtils.getDataByMap("customerData",2) );
        System.out.println("Mail Id -> " + ExcelUtils.getDataByMap("customerData",2).get("email_id") );
        System.out.println("Mail Id -> " + ExcelUtils.getDataByMap("customerData",1).get("email_id") );
    }
}
