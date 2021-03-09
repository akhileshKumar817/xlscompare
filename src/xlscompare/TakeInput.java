/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlscompare;

/**
 *
 * @author akhil
 */
public class TakeInput {
    String xlsA="";
    String xlsB="";
    String range="";
    String skiprange="";
    String xlsSheetA="";
    String xlsSheetB="";
    String resultFileName="";
    
    int [] ranges;

   public TakeInput(String xlsA, String xlsB, String range, String skiprange, String xlsSheetA,
				String xlsSheetB, String resultFileName) {
			super();
			this.xlsA = xlsA;
			this.xlsB = xlsB;
			this.range = range;
			this.skiprange = skiprange;
			this.xlsSheetA = xlsSheetA;
			this.xlsSheetB = xlsSheetB;
			this.resultFileName = resultFileName;
                        ranges=new int[4];
		}

    public String getXlsA() {
        return xlsA;
    }

    public String getXlsB() {
        return xlsB;
    }

    public String getRange() {
        
        return range;
    }

    public String getSkiprange() {
        return skiprange;
    }

    public String getXlsSheetA() {
        return xlsSheetA;
    }

    public String getXlsSheetB() {
        return xlsSheetB;
    }

    public String getResultFileName() {
        return resultFileName;
    }

    public int[] getRanges() {
        String rowStrart=range.split(":")[0].split("_")[1];
        String columnStart=range.split(":")[0].split("_")[0];
        
        String rowEnd=range.split(":")[1].split("_")[1];
        String columnEnd=range.split(":")[1].split("_")[0];
        
        ranges[0]=Integer.parseInt(rowStrart);
        ranges[1]=(int)columnStart.charAt(0) - (int)'a' + 1;
        ranges[2]=Integer.parseInt(rowEnd);
        ranges[3]=(int)columnEnd.charAt(0) - (int)'a' + 1;
        return ranges;
    }

   
    
    
}
