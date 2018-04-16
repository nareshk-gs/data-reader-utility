package com.datareader.utils;

import org.testng.annotations.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class ExcelTest {

    @Test
    public void loadExcel() {
        String file = System.getProperty("user.dir")
                + File.separator + "src" + File.separator + "test"
                + File.separator + "resources" + File.separator
                + "text.xlsx";

        List<Map<String, Object>> dataMap = new Excel().getDataAsList(new Excel().readExcel(file, "Sheet1"));
        for(Map<String, Object> map : dataMap) {
            System.out.println(map);
        }
    }
}
