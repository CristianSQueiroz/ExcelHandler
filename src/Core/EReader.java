/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Core;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

public class EReader {

    //private static final String FILE_NAME = "C:\\Users\\Cristian\\Downloads\\manda pro faro.xlsx";

    /*public static void main(String[] args) {

        try {
            printTime();
            boolean primeira = true;

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            MySqlConnect.getInstance().open();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                String cmd = "";

                while (cellIterator.hasNext()) {
                    
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        cmd += currentCell.getStringCellValue() + "--";
                        //System.out.print(currentCell.getStringCellValue() + "--");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        cmd += currentCell.getNumericCellValue() + "--";
                        //System.out.print(currentCell.getNumericCellValue() + "--");
                    }

                }
                if(primeira){
                    primeira = false;
                }else{
                    MySqlConnect.getInstance().executaComandoPadraoLote("CALL INSERTORUPDATE ('"+cmd+"')");
                }
                
                //System.out.println();

            }
            printTime();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            MySqlConnect.getInstance().close();
        }
        try {
            printTime();
            boolean primeira = true;

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            MySqlConnect.getInstance().open();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                String cmd = "INSERT INTO ATRIBUTO_STHT (NR_SEQUENCIA,VL_PROCURA,VL_REFERENCIA) VALUES (NULL,";

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        cmd += currentCell.getStringCellValue() + ",";
                        //System.out.print(currentCell.getStringCellValue() + "--");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        cmd += currentCell.getNumericCellValue() + ",";
                        //System.out.print(currentCell.getNumericCellValue() + "--");
                    }

                }
                cmd = cmd.substring(0, cmd.length() - 1);
                cmd = cmd + ");";
                System.err.println(cmd);
                MySqlConnect.getInstance().executaComandoPadrao(cmd);

                //System.out.println();
            }
            printTime();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            MySqlConnect.getInstance().close();
        }

    }*/
    
    private static void importOS(String filePath){
        try {
            printTime();
            boolean primeira = true;

            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            MySqlConnect.getInstance().open();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                String cmd = "";

                while (cellIterator.hasNext()) {
                    
                    Cell currentCell = cellIterator.next();
                    if (currentCell.getCellTypeEnum() == CellType.STRING) {
                        cmd += currentCell.getStringCellValue() + "--";
                        //System.out.print(currentCell.getStringCellValue() + "--");
                    } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                        cmd += currentCell.getNumericCellValue() + "--";
                        //System.out.print(currentCell.getNumericCellValue() + "--");
                    }

                }
                if(primeira){
                    primeira = false;
                }else{
                    MySqlConnect.getInstance().executaComandoPadraoLote("CALL INSERTORUPDATE ('"+cmd+"')");
                }
                
                //System.out.println();

            }
            printTime();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
            MySqlConnect.getInstance().close();
        }
    }

    public static void printTime() {
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy hh:mm:ss:ssss");
        System.out.println(sdf.format(new Date(System.currentTimeMillis())));
    }
}
