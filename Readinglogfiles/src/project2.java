import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Scanner;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class project2 {
	 static HashMap hmap = new HashMap();
     static XSSFWorkbook workbook = new XSSFWorkbook(); 
     static XSSFSheet spreadsheet = workbook.createSheet("Generated "+LocalDate.now());
     static XSSFRow row=spreadsheet.createRow(0);
     static XSSFCell cell;
     public static void main( String[] args ) throws IOException {
      File file =new File("D:/log/logging.log");
      String s[] = {"510229","510290","432651","427927","432658","432701"};
      Scanner in = null;
      try {
          in = new Scanner(file);
          while(in.hasNext())
          {
              String line=in.nextLine();
              for(String s1: s)
              {
                  if(line.contains(s1))
                  {
                	  String str[] = line.split("\\s");
                	  System.out.println(str[1]+"------"+s1);
                	  hmap.put(s1, str[1]);
                  }

              }
          }
      } catch (FileNotFoundException e) {
          e.printStackTrace();
      }
      
      cell=row.createCell(0);
      cell.setCellValue("Method Name");
      cell=row.createCell(1);
      cell.setCellValue("Executed Time");
      cell=row.createCell(2);
      cell.setCellValue("Execution Time");
      int i=1;
      Iterator entries =  hmap.entrySet().iterator();
      while(entries.hasNext())
      {
    	  Entry thisEntry = (Entry) entries.next();
    	  System.out.println(thisEntry.getKey()+"----"+thisEntry.getValue());
	      row=spreadsheet.createRow(i);
	         cell=row.createCell(0);
	         cell.setCellValue(thisEntry.getKey().toString());
	         cell=row.createCell(1);
	         cell.setCellValue(thisEntry.getValue().toString());
	         if(i%2==0)
	         {
	       	  cell = row.createCell(2);
	       	  String formula ="(A"+(i+1)+"-A"+i+")";
	    	  cell.setCellFormula(formula);
	         }
	         i++;
      }
      System.out.println(spreadsheet.getLastRowNum());
      FileOutputStream out = null;
			try {
				String exepath = "D://exceldata";
								String bufpath = exepath.concat(".xlsx");
								out = new FileOutputStream (new File(bufpath));
								workbook.write(out);
								System.out.println("Executed successfully 2");
//								SendFromYahoo.send();
								out.close();
						
			} catch (Exception e) {
				e.printStackTrace();
			}

		}

	}
