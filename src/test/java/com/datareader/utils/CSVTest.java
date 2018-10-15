package com.datareader.utils;

import org.testng.annotations.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

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
            if(map.get("div") == "001") {
                assertThat(map.get("store"), is(equalTo("00000")));
            }
    
            if(map.get("div") == "100") {
                assertThat(map.get("store"), is(equalTo("10000")));
            }
            
            if(map.get("div") == "003") {
                assertThat(map.get("sample"), is(equalTo("1,2,3")));
            }
        }
    }


}
