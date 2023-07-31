package org.example;
import org.utilKit.FileKit;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * 将格式为 文字——空白行——文字的txt文档存入excel
 *
 * by@hhb
 */
public class TextToExcelConverter {
    public static void main(String[] args) {
        FileKit kit=new FileKit();
        String filepath="D:\\FileTools\\files\\File_7_24_2\\";
        String inputFile = filepath+"input.txt";
        String outputFile = filepath+"object.xlsx";
        int targetColumn = 4; // 目标列的索引（从0开始）
        boolean thisHaveColor=true;
        kit.textToExcel(inputFile,outputFile,targetColumn,thisHaveColor);
    }
}
