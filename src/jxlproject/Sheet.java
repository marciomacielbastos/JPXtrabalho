/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package jxlproject;

import java.io.File;
import java.util.Date;
import jxl.read.biff.Formula;
import jxl.write.DateFormats;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WriteException;

/**
 *
 * @author marcio
 */
public class Sheet {
    private static void writeDataSheet(WritableSheet s) 
    throws WriteException
  {

    /* Formata a fonte */
    WritableFont wf = new WritableFont(WritableFont.ARIAL, 
      10, WritableFont.BOLD);
    WritableCellFormat cf = new WritableCellFormat(wf);
    cf.setWrap(true);

    /* Cria um label e escreve a data em uma célula da folha*/
    Label l = new Label(0,0,"Data",cf);
    s.addCell(l);
    WritableCellFormat cf1 = 
      new WritableCellFormat(DateFormats.FORMAT9);

    DateTime dt = 
      new DateTime(0,1,new Date(), cf1, DateTime.GMT);

    s.adCell(dt);
    
    /* Cria um label e escreve um float numver em uma célula da folha*/
    l = new Label(2,0,"Float", cf);
    s.addCell(l);
    WritableCellFormat cf2 = new WritableCellFormat(NumberFormats.FLOAT);
    Number n = new Number(2,1,3.1415926535,cf2) {};
    s.addCell(n);

    n = new Number(2,2,-3.1415926535, cf2) {};
    s.addCell(n);

    /* Cria um label e escreve um float number acima de 3 decimais 
       em uma célula da folha*/
    l = new Label(3,0,"3dps",cf);
    s.addCell(l);
    NumberFormat dp3 = new NumberFormat("#.###");
    WritableCellFormat dp3cell = new WritableCellFormat(dp3);
    n = new Number(3,1,3.1415926535,dp3cell);
    s.addCell(n);

    /* Cria um label e adiciona 2 células na folha*/
    l = new Label(4, 0, "Add 2 cells",cf);
    s.addCell(l);
    n = new Number(4,1,10);
    s.addCell(n);
    n = new Number(4,2,16);
    s.addCell(n);
    Formula f = new Formula(4,3, "E1+E2") {};
    s.addCell(f);

    /* Cria um Label e mulpiplica o valor de uma célula da folha por 2 */
    l = new Label(5,0, "Multiplica por 2",cf);
    s.addCell(l);
    n = new Number(5,1,10);
    s.addCell(n);
    f = new Formula(5,2, "F1 * 3");
    s.addCell(f);

    /* Cria um Label e divide o valor de uma célula da folha por 2.5 */
    l = new Label(6,0, "Divide por 2.5",cf);
    s.addCell(l);
    n = new Number(6,1, 12);
    s.addCell(n);
    f = new Formula(6,2, "F1/2.5");
    s.addCell(f);
  }

  private static void writeImageSheet(WritableSheet s) 
    throws WriteException
  {
    /* Cria um label e escreve uma imagem em uma célula da folha*/    
    Label l = new Label(0, 0, "Imagem");
    s.addCell(l);
    WritableImage wi = new WritableImage(0, 3, 5, 7, new File("imagem.png"));
    s.addImage(wi);

    /* Cria um label e escreve hyperlink em uma célula da folha*/
    l = new Label(0,15, "HYPERLINK");
    s.addCell(l);
    Formula f = new Formula(1, 15, 
      "DevMedia(\"http://www.devmedia.com.br\", "+
      "\"Portal DevMedia\")");
    s.addCell(f);
    
    }
}


