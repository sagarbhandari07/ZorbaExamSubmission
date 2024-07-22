package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteToExcel2 {
    public static void main(String[] args) throws IOException {
        String filePath = "C:\\Users\\sagar_quuljrl\\intije_Workspace\\Exam21July\\src\\main\\resources\\EmployeeTable.xlsx";

        FileInputStream fileInputStream = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        FileOutputStream fileOutputStream = new FileOutputStream(filePath);

        try {
            Sheet sheet = workbook.getSheet("Employees");
            String[] newColumns = {"manager_id", "emp_dept", "emp_share(%)"};


            Row headerRow = sheet.getRow(0);
            int colNum = headerRow.getLastCellNum();
            for (String newColumn : newColumns) {
                Cell newHeaderCell = headerRow.createCell(colNum++);
                newHeaderCell.setCellValue(newColumn);
            }
            Object[][] newData = {
                    {null, "Finance", 60},
                    {1001, "Finance", 20},
                    {1004, "R&D", 30},
                    {1004, "R&D", 40},
                    {1001, "Finance", 20},
                    {1005, "Finance", 15},
                    {1001, "Finance", 25}
            };

            int rowNum =1;
            for(Object[] rowData : newData) {
                Row row = sheet.getRow(rowNum++);
                if (row==null) {
                    row = sheet.createRow(rowNum -1);
                }
                colNum = headerRow.getLastCellNum();
                for (Object field : rowData) {
                    Cell cell = row.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Double) {
                        cell.setCellValue((Double) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    } else if (field == null) {
                        cell.setCellValue("");
                    }
                }
            }


            workbook.write(fileOutputStream);
            System.out.println("Data has been added to Excel file successfully!");

        } finally {

            if (fileOutputStream != null) {
                fileOutputStream.close();
            }
            if (workbook != null) {
                workbook.close();
            }
            if (fileInputStream != null) {
                fileInputStream.close();
            }
        }
    }

}





