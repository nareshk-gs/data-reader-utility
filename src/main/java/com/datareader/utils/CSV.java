package com.datareader.utils;

import com.opencsv.CSVReader;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.*;

public class CSV {

    private static CSVReader READER;


    public CSV() {

    }

    public CSV(String file) {
        READER = readFile(file);
    }

    public CSVReader readFile(String file)  {
        CSVReader reader = null;
        try{
            reader =  new CSVReader(new FileReader(file));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return reader;
    }

    public List<Map<String, String>> getDataAsList() {
        List<Map<String, String>> data = new ArrayList<>();
        List<List<String>> dataAsList = new ArrayList<>();
        String[] line;
        try{
            while((line = READER.readNext()) != null) {
                dataAsList.add(Arrays.asList(line));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        List<String> header = dataAsList.get(0);
        for(int i=1; i<dataAsList.size(); i++) {
            List<String> rowData = dataAsList.get(i);
            if (header.size() == rowData.size()) {
                Map<String, String> map = new HashMap<>();
                for(int j=0; j<header.size(); j++) {
                    map.put(header.get(j).trim(), rowData.get(j).trim());
                }
                data.add(map);
            }
        }
        return data;
    }
}
