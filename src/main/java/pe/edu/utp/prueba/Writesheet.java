
package pe.edu.utp.prueba;

import com.lowagie.text.Cell;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Writesheet {

    public static void main(String[] args)throws Exception {

        //create blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        
        //create a blank sheer
        XSSFSheet spreadsheet = workbook.createSheet("Employee Info");
        
        //Cerate row object
        XSSFRow row;

        //This data needs to be written (Object)

        Map < String, Object[] >empinfo =
        new TreeMap < String, Object[] >();
        empinfo.put("1", new Object[] {"EMP ID", "EMP NAME", "DESIGNATION"});
        empinfo.put("2", new Object[] {"tp01", "Gopal", "Thecnical Manager"});
        empinfo.put("3", new Object[] {"tp02", "Manisha", "Proof reader"});
        empinfo.put("4", new Object[] {"tp03", "Masthan", "Technical writer"});
        empinfo.put("5", new Object[] {"tp04", "Satish", "Technical writer"});
        empinfo.put("6", new Object[] {"tp05", "Krisha", "Technical writer"});
        
        //Iterate over data and write to shee
        Set < String > keyid = empinfo.keySet();
        int rowid = 0;
        
        for(String key : keyid){
            row = spreadsheet.createRow(rowid++);
            Object [] objectArr = empinfo.get(key);
            int cellid =0;
            
            for(Object obj : objectArr){
                XSSFCell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
            
        //Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File("Writessheet.xlsx"));
        workbook.write(out);
        out.close();
        System.out.println("Writesheet.xslx written succesfully");

    }
    
}
