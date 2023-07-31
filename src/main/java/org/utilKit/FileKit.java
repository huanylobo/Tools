package org.utilKit;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.example.ExcelColumnToTxt;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

/**
 * @author hhb
 * @date 2023/7/31
 */
public class FileKit {
    public static void textToExcel(String inputFile,String outputFile,int targetColumn,boolean thisHaveColor){
        FileKit kit=new FileKit();
        ArrayList<String> fields = new ArrayList<>();
        ExcelColumnToTxt ExToTxt=new ExcelColumnToTxt();
        try {
            // 读取文本文件
            FileInputStream file = new FileInputStream(outputFile);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            // 找到目标列的最后一个空白单元格的行号
            int lastRowNum = 0;
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(targetColumn);
                boolean havecolor=kit.haveColor(cell);
//                if ((cell == null || cell.getCellType() == CellType.BLANK||"".equals(cell.getStringCellValue()))
//                        ||((cell == null || cell.getCellType() == CellType.BLANK||"".equals(cell.getStringCellValue()))
//                        &&havecolor&&thisHaveColor))
                if(cell == null || cell.getCellType() == CellType.BLANK||"".equals(cell.getStringCellValue()))
                {
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
            lastRowNum=0;
            // 将字段写入到xlsx文件的指定列的空白单元格中
            for (String field : fields) {
                Row row = sheet.getRow(lastRowNum);
                if (row == null) {
                    row = sheet.createRow(lastRowNum);
                }
                Cell cell = row.getCell(targetColumn);

                if ( cell == null ||cell.getCellType() == CellType.BLANK||"".equals(cell.getStringCellValue())) {
                    cell = row.createCell(targetColumn);
                    cell.setCellValue(field);
                    lastRowNum++;
                } else {
                    lastRowNum++;
                    while ( row != null &&
                            row.getCell(targetColumn) != null &&
                            row.getCell(targetColumn).getCellType() != CellType.BLANK&&
                            !"".equals(row.getCell(targetColumn).getStringCellValue())) {
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
            System.out.println("长度："+lastRowNum);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    // 将列数据保存为txt文件
    public static void saveColumnDataToFile(String columnData, String filePath) throws IOException {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(filePath))) {
            writer.write(columnData);
        }
    }

    /**
     * columnIndex 为需要读取（翻译）的目标列
     * columnIndex1 为需要填入的空白列
     * colorCondition 为是否判断颜色
     * filePath 文件路径
     * @param
     * @return
     */
    public static void excelToTxt(int columnIndex,int columnIndex1,boolean colorCondition,String filePath){

        try {
            // 读取Excel文件
            Workbook workbook = new XSSFWorkbook(filePath+"object.xlsx");

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);
            // 创建一个用于保存列数据的StringBuilder
            StringBuilder columnData = new StringBuilder();
            int times = 0;
            // 遍历每一行并获取指定列的数据
            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                Cell cell1 = row.getCell(columnIndex1);
                // 获取第二列字符串
//                String cellValue1 = cell1.getStringCellValue();

                if (cell1==null||cell1.getCellType() == CellType.BLANK || "".equals(cell1.getStringCellValue()))
//                if((cell1.getCellType() == CellType.BLANK || "".equals(cellValue1))||((cell1.getCellType() == CellType.BLANK || "".equals(cellValue1))&&colorCondition&&haveColor(cell)))
                {
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
            String outputFilePath = filePath+"output.txt";
            saveColumnDataToFile(columnData.toString(), outputFilePath);
            System.out.println("cishu: "+times);
            System.out.println("数据保存完成！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static boolean haveColor(Cell cell){
        CellStyle cellStyle = cell.getCellStyle();
        short foregroundColorIndex = cellStyle.getFillForegroundColor();
        Color foregroundColor = cellStyle.getFillForegroundColorColor();
        boolean hasColor = (foregroundColorIndex != 0) || (foregroundColor != null);
        return hasColor;
    }
}
