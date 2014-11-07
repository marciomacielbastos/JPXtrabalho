/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jxlproject;
import java.io.File;
import java.io.IOException;
import jxl.Cell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.write.Number;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;


/**
 *
 * @author marcio
 */
public class JXLproject {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        try{
            WritableWorkbook wworkbook;
            wworkbook = Workbook.createWorkbook(new File("JXLproject.xls"));
            WritableSheet wsheet = wworkbook.createSheet("", 0);
            Label label = new Label(0, 2, "A label record");
            wsheet.addCell(label);
            Number number = new Number(3, 4, 3.1459) {};
            wsheet.addCell(number);
            wworkbook.write();
            Workbook workbook = Workbook.getWorkbook(new File("/home/marcio/NetBeansProjects/JXLprojecT/PlanilhaOrcamentoPessoalEdu.xls.xls"));
            Sheet sheet = workbook.getSheet(0);
            Cell cell1 = sheet.getCell(0, 2);
            System.out.println(cell1.getContents());
            Cell cell2 = sheet.getCell(3, 4);
            System.out.println(cell2.getContents());
            workbook.close();
            
       }
       catch(IOException ioe){
           System.out.println(ioe.getMessage());
       }
       catch(BiffException be){
           System.out.println(be.getMessage());
       }
       catch(WriteException we){
           System.out.println(we.getMessage());
       }
    }
    
}
