/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlscompare;

import java.util.ArrayList;
import java.util.Arrays;

/**
 *
 * @author akhil
 */
public class Test {
    
    public static void main(String[] args) {
       /* TakeInput t=new TakeInput("","","a_1:d_7","","","","");
        
        
        int test1[]=t.getRanges();
        
        ReadXLS r=new ReadXLS("C:\\Users\\akhil\\OneDrive\\Desktop\\Test\\Book1.xlsx");
       
       ArrayList<String[]> s=r.getDataByRange(test1[0]-1, test1[1]-1, test1[2]-1, test1[3]-1);
       for(String[] a:s)
       {
           System.out.println(Arrays.toString(a));
       }
       */
       
       
       Compare c=new Compare();
       c.compareResult();//for Testing application Tested and added 1.0
       
    }
}
