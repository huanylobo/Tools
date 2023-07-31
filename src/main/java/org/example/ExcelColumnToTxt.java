package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.utilKit.FileKit;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;

public class ExcelColumnToTxt {
    public static void main(String[] args) {
        FileKit kit=new FileKit();
        int columnIndex=1;
        int columnIndex1=4;
        boolean colorCondition =true;
        String filePath="D:\\FileTools\\files\\File_7_24_2\\";
        FileKit.excelToTxt(columnIndex,columnIndex1,colorCondition,filePath);

    }


}
