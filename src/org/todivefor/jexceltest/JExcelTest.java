/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.todivefor.jexceltest;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.write.DateTime;
import jxl.write.WritableCellFormat;
import jxl.write.WriteException;


/**
 *
 * @author peterream
 */
public class JExcelTest
{

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args)
    {
        try
        {
            WritableWorkbook workbook = Workbook.createWorkbook(new File("JExceltest.xls"));
            WritableCellFormat dateCellFormatMDY = new WritableCellFormat
                (new jxl.write.DateFormat("mm/dd/yy"));                     // Used to format date
            
            WritableSheet sheet = workbook.createSheet("First Sheet", 0);
            int col = 0;
            int row = 2;
            Label label = new Label(col, row, "A label record");
            row = 3;
            Label label1 = new Label(col, row, "A 2nd label record");
            sheet.addCell(label);
            sheet.addCell(label1);
            col = 0;
            row = 4;
            for (col = 1; col < 5; col++)
            {
                Number number = new Number(col, row, col);
                sheet.addCell(number); 
            }
            String myDate = "September 11, 2001";                       // String date
            DateFormat df = new SimpleDateFormat("MMM dd, yyyy");
            Date startDate = null;
            try
            {
                startDate = df.parse(myDate);
                System.out.println(startDate);
            }
            catch (ParseException ex)
            {
                Logger.getLogger(JExcelTest.class.getName()).log(Level.SEVERE, null, ex);
            }
            col = 0;
            row = 10;
            DateTime date = new DateTime(col, row, startDate, dateCellFormatMDY);
            sheet.addCell(date);
            workbook.write();
            workbook.close();
        }
        catch (IOException ex)
        {
            Logger.getLogger(JExcelTest.class.getName()).log(Level.SEVERE, null, ex);
        }
        catch (WriteException ex)
        {
            Logger.getLogger(JExcelTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
