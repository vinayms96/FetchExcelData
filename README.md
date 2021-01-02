# FetchExcelData
This is the script to fetch the row data from Excel File (.xlsx) format in two ways.
1. Get the entire row data by passing the **Sheet Name** & **Data Name** as argument and adding the entire row data to an ArrayList.

    ```public static ArrayList<String> getDataByList(String sheetName, String dataName)```
    
2. Get the entire row data by passing the **Sheet Name** & **Row Number** for which the data is required, and adding it to a HashMap 

    ```public static HashMap<String, String> getDataByMap(String sheetName, int rowNum```