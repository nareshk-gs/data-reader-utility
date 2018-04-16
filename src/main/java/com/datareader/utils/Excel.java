package com.datareader.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by KON1299 on 6/28/2017.
 */
public class Excel {

    private XSSFWorkbook book;
    private XSSFSheet dataSheet;
    private XSSFRow row;
    private XSSFCell cell;

    /**
     * Reads the sheet from a defined file and returns the sheet as an instance of XSSFSheet
     * @param file - absolute path of the file, including file name
     * @param sheetName - name of the sheet to be read
     * @return Sheet as an instance of XSSFSheet
     */

    public XSSFSheet readExcel(String file, String sheetName){

        try{
            FileInputStream excelFile = new FileInputStream(new File(file));
            book = new XSSFWorkbook(excelFile);

        }catch (FileNotFoundException fnfe){
            System.out.println(fnfe.getMessage());
        }catch (IOException ioe){
            ioe.getMessage();
        }
        return dataSheet = book.getSheet(sheetName);
    }


    /**
     * Returns number of rows exist in the sheet
     * @param sheet Sheet to be read, as an instance of XSSFSheet
     * @return number of rows in the sheet
     */
    public int getRowCount(XSSFSheet sheet){
        return sheet.getLastRowNum();
    }


    /**
     * Returns the given row as an instance of XSSFRow
     * @param sheet Sheet name to be read from
     * @param rowNum row umber to be read
     * @return Data on the row as an instance of XSSFRow
     */
    public XSSFRow readRow(XSSFSheet sheet, int rowNum){
        return sheet.getRow(rowNum);
    }


    /**
     * Returns number of columns exist in the sheet
     * @param row row to be read for number of columns, as an instance of XSSFRow
     * @return number of columns in the row
     */
    public int getColumnCount(XSSFRow row){
        return row.getPhysicalNumberOfCells();
    }


    /**
     * Returns data in the given row, as an instance of XSSFRow, in an ArrayList
     * @param row row to be read, as an instance of XSSFRow
     * @return All the data in row as elements in ArrayList
     */
    public ArrayList<Object> readRowData(XSSFRow row){
        int columnCount=0;
        if(row!=null){
            columnCount = getColumnCount(row);
        }

        ArrayList<Object> rowData = new ArrayList();
        for(int i=0;i<columnCount;i++){
            XSSFCell cell = row.getCell(i);
            if(cell.getCellTypeEnum()== CellType.STRING){
                //System.out.println(cell.getStringCellValue());
                rowData.add(cell.getStringCellValue());
            }
            if(cell.getCellTypeEnum()==CellType.NUMERIC){
                //System.out.println(cell.getNumericCellValue());
                rowData.add(cell.getNumericCellValue());
            }
        }
        System.out.println(rowData.size());
        return rowData;

    }


    /**
     * Returns entire sheet as a list of maps, with each element in list refers to a row on the sheet and
     * map contains the column headers as keys and cell values in that row as value
     * @param sheet sheet to be read, as an instance of XSSFSheet
     * @return List of Map of String and Object
     */
    public List<Map<String, Object>> getDataAsList(XSSFSheet sheet){
        List<Map<String, Object>> sheetData = new ArrayList();
        int rowCount = sheet.getLastRowNum();
        if(rowCount>0){
            ArrayList<String> header = new ArrayList();
            row = sheet.getRow(0);
            int colCount = row.getPhysicalNumberOfCells();
            for(int i=0;i<colCount;i++){
                cell = row.getCell(i);
                if(cell.getCellTypeEnum()== CellType.STRING){
                    header.add(cell.getStringCellValue());
                }
            }

            //System.out.println(header);

            for(int i=1;i<=rowCount;i++){
                Map<String, Object> map = new HashMap<>();
                row = sheet.getRow(i);
                for(int j=0;j<header.size();j++){
                    cell = row.getCell(j);
                    if(cell == null){
                        map.put(header.get(j),null);
                    }else{
                        if(cell.getCellTypeEnum()== CellType._NONE){
                            map.put(header.get(j),null);
                        }
                        if(cell.getCellTypeEnum()== CellType.STRING){
                            map.put(header.get(j),cell.getStringCellValue());
                        }
                        if(cell.getCellTypeEnum() == CellType.NUMERIC){
                            map.put(header.get(j),cell.getNumericCellValue());
                        }
                        if(cell.getCellTypeEnum() == CellType.BLANK){
                            map.put(header.get(j),"");
                        }
                    }
                }
                sheetData.add(map);
            }
        }
        return sheetData;
    }

    /**
     * Creates a file with given name in the test/resources/results directory of the current project
     * file name would be in the format <given test name>_<date in yyyyMMddHHmm format>.xlsx
     * @param testName
     * @return name of the file created
     * @throws FileNotFoundException
     */
    public String createFile(String testName) throws FileNotFoundException{
        String filePath =
                System.getProperty("user.dir")
                + File.separator + "src" + File.separator + "test" + File.separator + "resources"
                + File.separator + "results";
        File resultsFolder = new File(filePath);
        if(!resultsFolder.exists()){
            resultsFolder.mkdir();
        }
        String date = new SimpleDateFormat("yyyyMMddHHmm").format(new Date());

        String fileName = testName+"_"+date+".xlsx";
        String vfile = filePath+File.separator+fileName;
        System.out.println(vfile);
        new FileOutputStream(new File(vfile));
        return vfile;
    }


    /**
     * Creates a row and writes data from a given ArrayList into a given sheet in the given file
     * @param fileName
     * @param sheetName
     * @param rowList
     * @throws IOException
     * @throws InvalidFormatException
     */
    public void writeRow(String fileName,String sheetName, ArrayList rowList) throws IOException,InvalidFormatException{
        File oFile = new File(fileName);
        FileInputStream input = new FileInputStream(oFile);
        book = new XSSFWorkbook(input);
        dataSheet = book.getSheet(sheetName);
        //Sheet sheet = workbook.getSheetAt(0);

        int rowNum = dataSheet.getLastRowNum();
        System.out.println("Last Row: "+rowNum);
        XSSFRow row = dataSheet.createRow(++rowNum);
        System.out.println("New Row: "+row);
        int cellNumber = 0;
        for(Object cellVal : rowList){
            System.out.println("Current Cell: "+ cellNumber);
            XSSFCell cell = row.createCell(cellNumber);
            System.out.println(cell+"-"+cellNumber);
            if(cellVal instanceof String){
               cell.setCellValue((String) cellVal);
            } else {
                cell.setCellValue("String");
            }
            cellNumber++;
        }
        input.close();

        FileOutputStream fos = new FileOutputStream(oFile);
        book.write(fos);
        book.close();
        fos.close();


    }


    /**
     * Creates a sheet in the given file, create first row in the sheet and write data to the row from ArrayList.
     *
     * @param fileName
     * @param sheetName
     * @param headerList
     * @return file name as a string
     * @throws IOException
     */
    public String writeHeader(String fileName, String sheetName, List<String> headerList) throws IOException{
        FileOutputStream fos = new FileOutputStream(new File(fileName));
        XSSFWorkbook book = new XSSFWorkbook();
        XSSFSheet sheet = book.createSheet(sheetName);
        Row row = sheet.createRow(0);
        int cellNumber = 0;
        for(Object header : headerList){
            Cell cell = row.createCell(cellNumber++);
            cell.setCellValue((String) header);
        }
        book.write(fos);
        book.close();
        fos.close();
        return fileName;
    }
}
