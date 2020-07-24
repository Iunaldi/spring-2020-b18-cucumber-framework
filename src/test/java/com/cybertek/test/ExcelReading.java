package com.cybertek.test;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ExcelReading {

    public static void main(String[] args) {


        try (FileInputStream fileInputStream = new FileInputStream("src/SampleData.xlsx")) {
            XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheet("Employees");
            System.out.println(sheet.getRow(2).getCell(1));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
