package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;

/**
 * 可用的提取excel程序
 * 用于提取位于某一列有而另一列无的格子
 *
 * by@hhb
 */
public class ExcelColumnToTxt2 {
    public static void main(String[] args) {
        try {
            // 读取Excel文件
            Workbook workbook = new XSSFWorkbook("D:\\FileTools\\files\\File_6_5\\LangueConfig_test.xlsx");

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);

            // 获取需要读取的列的索引
            int columnIndex = 3;
            // 筛选用的列的索引
            int columnIndex1 = 2;
            // 创建一个用于保存列数据的StringBuilder
            StringBuilder columnData = new StringBuilder();
            int times = 0;
            // 遍历每一行并获取指定列的数据
            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                Cell cell1 = row.getCell(columnIndex1);
                // 获取第二列字符串
                String cellValue1 = null;

                if (cell1 == null || cell1.getCellType() == CellType.BLANK) {
                    if (cell != null) {
                        String cellValue = "";
                        // 获取第三列字符串
                        if (cell.getCellType() == CellType.STRING) {
                            cellValue = cell.getStringCellValue();
                        } else if (cell.getCellType() == CellType.NUMERIC) {
                            cellValue = String.valueOf(cell.getNumericCellValue());
                        }

                        columnData.append(cellValue).append("\n");
                    }
                    columnData.append("\n");
                    times++;
                }
            }

            // 将列数据保存为txt文件
            String outputFilePath = "D:\\FileTools\\files\\File_6_5\\output.txt";
            saveColumnDataToFile(columnData.toString(), outputFilePath);

            System.out.println("数据保存完成！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 将列数据保存为txt文件
    private static void saveColumnDataToFile(String columnData, String filePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            writer.write(columnData);
        }
    }

    //判断单元格是否有颜色
    private static boolean haveColor(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();
        short foregroundColorIndex = cellStyle.getFillForegroundColor();
        Color foregroundColor = cellStyle.getFillForegroundColorColor();
        boolean hasColor = (foregroundColorIndex != 0) || (foregroundColor != null);
        return hasColor;
    }
}
