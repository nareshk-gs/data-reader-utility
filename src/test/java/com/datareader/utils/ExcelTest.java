package com.datareader.utils;

import org.testng.annotations.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;

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
            if(map.get("Sl No") == "1") {
                assertThat(map.get("Header1"), is(equalTo("test11")));
            }
    
            if(map.get("Sl No") == "2.0") {
                assertThat(map.get("Header1"), is(equalTo("test21")));
            }
        }
    }
}
