/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlscompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author akhil
 */
public class ReadXLS {

    String xlsPath;
    int sheetIndex = 0;

    File currentFile;
    XSSFSheet currentSheet;

    public ReadXLS(String xlsPath, int sheetIndex) {
        this.xlsPath = xlsPath;
        this.sheetIndex = sheetIndex;
        iniFileObject();
        iniSheet();
    }

    public ReadXLS(String xlsPath) {
        this.xlsPath = xlsPath;
        iniFileObject();
        iniSheet();
    }

    public void iniFileObject() {
        try {
            this.currentFile = new File(xlsPath);
        } catch (Exception e) {
            System.out.println(e.toString());
        }
    }

    public void iniSheet() {
        try {
            FileInputStream inputStream = new FileInputStream(currentFile);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            currentSheet = workbook.getSheetAt(sheetIndex);
        } catch (IOException ex) {
            System.out.println(ex.toString());
        }
    }

    public ArrayList<String[]> getDataByRange(int rowStrart, int columnStart, int rowEnd, int columnEnd) {
        ArrayList<String[]> result = new ArrayList();

        for (int i = rowStrart; i <= rowEnd; i++) {
            XSSFRow currentRow = currentSheet.getRow(i);
            String temp[] = new String[columnEnd + 1];
            for (int j = columnStart; j <= columnEnd; j++) {
                XSSFCell currentCell = currentRow.getCell(j);
                try {
                    String s = currentCell.getStringCellValue();
                    temp[j] = s+"_"+currentCell.getReference();
                   // System.out.println(currentCell.getReference());
                } catch (java.lang.NullPointerException ex) {
                    currentRow.createCell(j).setCellValue("");
                    currentCell = currentRow.getCell(j);
                    temp[j] = ""+"_"+currentCell.getReference();
                    
                } catch (java.lang.IllegalStateException ex) {
                   // System.out.println(ex.toString());
                    String s = String.valueOf(currentCell.getNumericCellValue());

                    if (currentCell.getCellType() == currentCell.getCellType().NUMERIC && HSSFDateUtil.isCellDateFormatted(currentCell)) {
                        Date javaDate = DateUtil.getJavaDate(currentCell.getNumericCellValue());
                        //System.out.println(new SimpleDateFormat("MM/dd/yyyy").format(javaDate));
                         temp[j] = new SimpleDateFormat("MM/dd/yyyy").format(javaDate)+"_"+currentCell.getReference();
                    }
                    else
                    {
                         temp[j] = s+"_"+currentCell.getReference();
                    }

                   
                } catch (Exception e) {
                    System.out.println(e.toString());
                }
                //  System.out.println(s);

            }
            result.add(temp);
        }

        return result;
    }
}
