package org.scott.tests;
 
import java.util.List;

import org.testng.annotations.*;
 
public class MyTests {

    @DataProvider(name = "ExcelSheet")
    public static Object[][] ExcelSheet() {
        return ExcelUtil.readSheet("src/main/resources/testdata.xlsx");
    }
     
    @Test(dataProvider = "ExcelSheet")
    public void Test1(List<String> columns) {
        System.out.println("Testing row with " + columns.size() + " columns");
        System.out.println("Column 0 = " + columns.get(0));
    }
}
