package org.example;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelColumnToTxt {
    public static void main(String[] args) {
        try {
            // 读取Excel文件
            Workbook workbook = new XSSFWorkbook("D:\\FileTools\\files\\Language1.xlsx");

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);

            // 获取需要读取的列的索引
            int columnIndex = 3;
            // 筛选用的列的索引
            int columnIndex1 = 2;
            // 创建一个用于保存列数据的StringBuilder
            StringBuilder columnData = new StringBuilder();
            int times=0;
            // 遍历每一行并获取指定列的数据
            for (Row row : sheet) {

                Cell cell = row.getCell(columnIndex);
                Cell cell1 = row.getCell(columnIndex1);
                //获取第二列字符串
                String cellValue1 = null;
                //判断条件1
                if (cell1 == null) {
                    times++;
                    continue;
                }
                if (cell1.getCellType() == CellType.STRING) {
                    cellValue1 = cell.getStringCellValue();
                } else if (cell1.getCellType() == CellType.NUMERIC) {
                    cellValue1 = String.valueOf(cell.getNumericCellValue());
                }
                //判断条件2
                if(cellValue1 != null){
                    times++;
                    continue;
                }
                    if (cell != null) {
                        String cellValue = "";
                        //获取第三列字符串
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
//            }

            // 将列数据保存为txt文件
            String outputFilePath = "D:\\FileTools\\files\\output.txt";
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
}
