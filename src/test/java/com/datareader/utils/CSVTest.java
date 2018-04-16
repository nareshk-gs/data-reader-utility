package com.datareader.utils;

import org.testng.annotations.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class CSVTest {

    @Test
    public void loadCsvTest() {
        String file = System.getProperty("user.dir")
                + File.separator + "src" + File.separator + "test"
                + File.separator + "resources" + File.separator
                + "test.csv";

        List<Map<String, String>> dataMap = new CSV(file).getDataAsList();
        for(Map<String, String> map : dataMap) {
            System.out.println(map);
        }
    }


}
