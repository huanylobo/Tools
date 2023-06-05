package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 将格式为 文字——空白行——文字的txt文档存入excel
 *
 * by@hhb
 */
public class TextToExcelConverter {
    public static void main(String[] args) {
        String inputFile = "D:\\FileTools\\files\\File_6_5\\input.txt";
        String outputFile = "D:\\FileTools\\files\\File_6_5\\LangueConfig_test.xlsx";
        int targetColumn = 2; // 目标列的索引（从0开始）

        ArrayList<String> fields = new ArrayList<>();

        try {
            // 读取文本文件
            FileInputStream file = new FileInputStream(outputFile);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            // 找到目标列的最后一个空白单元格的行号
            int lastRowNum = sheet.getLastRowNum();
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(targetColumn);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    lastRowNum = row.getRowNum();
                    break;
                }
            }

            // 读取文本文件并将字段添加到ArrayList中
            BufferedReader reader = new BufferedReader(new FileReader(inputFile));
            String line;
            StringBuilder fieldBuilder = new StringBuilder();
            while ((line = reader.readLine()) != null) {
                line = line.trim();
                if (!line.isEmpty()) {
                    fieldBuilder.append(line).append("\n");
                } else {
                    if (fieldBuilder.length() > 0) {
                        String field = fieldBuilder.toString().trim();
                        fields.add(field);
                        fieldBuilder.setLength(0); // 重置StringBuilder
                    }
                }
            }
            reader.close();

            // 检查是否还有剩余的字段
            if (fieldBuilder.length() > 0) {
                String field = fieldBuilder.toString().trim();
                fields.add(field);
            }

            // 将字段写入到xlsx文件的指定列的空白单元格中
            for (String field : fields) {
                Row row = sheet.getRow(lastRowNum);
                if (row == null) {
                    row = sheet.createRow(lastRowNum);
                }
                Cell cell = row.getCell(targetColumn);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    cell = row.createCell(targetColumn);
                    cell.setCellValue(field);
                    lastRowNum++;
                } else {
                    lastRowNum++;
                    while (row != null && row.getCell(targetColumn) != null &&
                            row.getCell(targetColumn).getCellType() != CellType.BLANK) {
                        row = sheet.getRow(lastRowNum);
                        lastRowNum++;
                    }
                    if (row == null) {
                        row = sheet.createRow(lastRowNum - 1);
                    }
                    cell = row.createCell(targetColumn);
                    cell.setCellValue(field);
                }
            }

            // 保存工作簿到文件
            FileOutputStream outputStream = new FileOutputStream(outputFile);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

            System.out.println("字段已成功写入Excel文件。");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
