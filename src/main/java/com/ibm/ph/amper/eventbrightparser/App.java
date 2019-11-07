package com.ibm.ph.amper.eventbrightparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Pro
 *
 */
public class App 
{
    private static String path = "C:\\Users\\CharlesAMPER\\Desktop\\joy\\report-2019-11-07T1138.xlsx";
    private static final String FILE_NAME = "C:\\Users\\CharlesAMPER\\Desktop\\joy\\report-2019-11-07T1138_output_1.xlsx";
    public static void main( String[] args )
    {
        
        XSSFWorkbook workbookout = new XSSFWorkbook();
        XSSFSheet sheet = workbookout.createSheet("New sheet");
        Workbook workbook = null;

        try {
            
            
            FileInputStream excelFile = new FileInputStream(new File(path));
            workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            int rowCount = 0;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Row row = sheet.createRow(rowCount++);
                String fName = "";
                String lName = "";
                String email = "";
                String phone = "";

                try {
                    Cell cell1=currentRow.getCell(4);
                    fName = cell1.getStringCellValue();
                } catch (Exception e) {
                    // TODO: handle exception
                }
                
                try {
                    Cell cell2=currentRow.getCell(5);
                    lName = cell2.getStringCellValue();
                    
                }  catch (Exception e) {
                    // TODO: handle exception
                }
                
                try {
                    Cell cell3=currentRow.getCell(6);
                    email = cell3.getStringCellValue();
                    
                }  catch (Exception e) {
                    // TODO: handle exception
                }
                
                try {
                    Cell cell15=currentRow.getCell(15);
                    phone = cell15.getStringCellValue();
                }  catch (Exception e) {
                    // TODO: handle exception
                }
                
                String name = fName + " " + lName;
                
                Cell cell_1 = row.createCell(0);
                Cell cell_2 = row.createCell(1);
                Cell cell_3 = row.createCell(2);
                
                cell_1.setCellValue(name);
                cell_2.setCellValue(email);
                cell_3.setCellValue(phone);
                
                System.out.println(phone);
            }    
        } catch (Exception ex) {
            System.out.println(ex);
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
        
        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbookout.write(outputStream);
            workbookout.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
