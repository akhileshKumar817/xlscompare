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
public class Compare {
    
    
    
    
    public void compareResult()
    {
        ArrayList <String[]>compareR=new ArrayList();
        TakeInput input=new TakeInput("C:\\Users\\akhil\\OneDrive\\Desktop\\Test\\Book1.xlsx","C:\\Users\\akhil\\OneDrive\\Desktop\\Test\\Book2.xlsx","a_1:d_7","","","","");
        int test1[]=input.getRanges();
        
        ReadXLS r=new ReadXLS(input.getXlsA());
       ArrayList<String[]> s=r.getDataByRange(test1[0]-1, test1[1]-1, test1[2]-1, test1[3]-1);
       
        ReadXLS r1=new ReadXLS(input.getXlsB());
       ArrayList<String[]> s1=r1.getDataByRange(test1[0]-1, test1[1]-1, test1[2]-1, test1[3]-1);
       
       String dataResult[];
       for(int i=0;i<s.size();i++)
       {
           String dataA[]=s.get(i);
           String dataB[]=s1.get(i);
           dataResult=new String [dataA.length];
           for(int j=0;j<dataA.length;j++)
           {
               if(dataA[j].equals(dataB[j]))
               {
                   dataResult[j]=dataA[j]+"_"+dataB[j]+"_"+"Matched";
               }
               else
               {
                    dataResult[j]=dataA[j]+"_"+dataB[j]+"_"+"NotMatched";
               }
               
           }
           
           compareR.add(dataResult);
       }
        
        
       for(String a[]:compareR)
       {
           System.out.println(Arrays.toString(a));
       }
    }
}
