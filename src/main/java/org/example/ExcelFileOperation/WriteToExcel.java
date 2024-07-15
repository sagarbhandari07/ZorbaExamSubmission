package org.example.ExcelFileOperation;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import javax.xml.crypto.Data;

public class WriteToExcel {
    public static void main(String[] args) {
        String excelFilePath = "C:\\Users\\sagar_quuljrl\\intije_Workspace\\Exam15Jul2024\\src\\main\\resources\\EmployeeTable.xlsx";
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Employees");


            Object[][] employees = {
                    {"Employee_id", "Employee_name", "Emp_Salary", "Emp_mobile", "Emp_city"},
                    {1001, "Jack", 1482.45, "0809808008", "NYC"},
                    {1002, "Joy", 5282.12, "9809808008", "SD"},
                    {1003, "Nick", 3454.11, "8976876786", "Dayton"},
                    {1004, "Joe", 6482.45, "8809808008", "NYC"},
                    {1005, "Nick", 5482.45, "5809808008", "CA"},
                    {1006, "Hyder", 9482.45, "2809808008", "LA"},
                    {1007, "Harry", 1182.45, "4809808008", "Ohio"}
            };

            int rowNum = 0;
            for (Object[] rowData : employees) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (Object field : rowData) {
                    Cell cell = row.createCell(colNum++);
                    if (field instanceof String) {
                        cell.setCellValue((String) field);
                    } else if (field instanceof Double) {
                        cell.setCellValue((Double) field);
                    } else if (field instanceof Integer) {
                        cell.setCellValue((Integer) field);
                    }
                }
            }


            try (FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOut);
                System.out.println("Excel file has been created successfully!");
            }

        }
        catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}