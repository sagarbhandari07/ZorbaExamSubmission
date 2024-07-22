package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

public class ReadFromExcel {
    public static void main(String[] args) throws IOException {
        String excelFilePath = "C:\\Users\\sagar_quuljrl\\intije_Workspace\\Exam15Jul2024\\src\\main\\resources\\EmployeeTable.xlsx";

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
             Scanner scanner = new Scanner(System.in)) {

            Sheet sheet = workbook.getSheet("Employees");


            System.out.print("Enter employee_id to search: ");
            int employeeId = scanner.nextInt();


            boolean get = false;
            for (Row row : sheet) {
                    int id = (int) idCell.getNumericCellValue();
                    if (id == employeeId) {

                        get = true;
                        System.out.println("Employee Details:");
                        System.out.println("Employee_id: " + id);
                        System.out.println("Employee_name: " + getStringValue(row.getCell(1)));
                        System.out.println("Emp_Salary: " + getNumericValue(row.getCell(2)));
                        System.out.println("Emp_mobile: " + getStringValue(row.getCell(3)));
                        System.out.println("Emp_city: " + getStringValue(row.getCell(4)));
                        System.out.println("Manager_id: " + getNumericValue(row.getCell(5)));
                        System.out.println("Emp_dept: " + getStringValue(row.getCell(6)));
                        System.out.println("Emp_share (%): " + getNumericValue(row.getCell(7)));
                        break;
                    }
                }
            }



        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private static String getStringValue(Cell cell) {
        return cell;
    }


    private static double getNumericValue(Cell cell) {
         return cell;
    }





