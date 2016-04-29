package com.company;

//import org.apache.poi.x
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException{
	// write your code here

        FileInputStream readStream = new FileInputStream("hello_poi_hssf.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(readStream);

        System.out.println(workbook.getNumberOfSheets());

        HSSFSheet sheet = workbook.getSheetAt(0);

        for (int r = 0 ; r < sheet.getPhysicalNumberOfRows() ; r++) {

            System.out.println(sheet.getRow(r));
            HSSFRow row = sheet.getRow(r);
            for (int c = 0 ; c < row.getPhysicalNumberOfCells() ; c++) {
                System.out.println(row.getCell(c));
                }

        }


    }
}
