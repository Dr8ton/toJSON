package xlstojson;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class Main {

    public static void main(String[] args) throws FileNotFoundException, IOException, InvalidFormatException {
        String path ="/home/dr8ton/Projects/postla.xlsx";
        File file = new File(path); 
               Workbook wb = WorkbookFactory.create(file);
            
        System.out.println("{  \"posts\" : {  " );

        for (Sheet sheet : wb ) {
        for (Row row : sheet) {
            
            Cell name = row.getCell(0);
            Cell address = row.getCell(1);
            Cell city = row.getCell(2);
            Cell state = row.getCell(3);
            Cell zipCode = row.getCell(4);
            Cell latitude = row.getCell(5);
            Cell longitude = row.getCell(6);
            Cell code = row.getCell(7);
            
     
            System.out.println("    \"" +name+ "\" : {\n" +
"        \"name\": \"" +name+ "\", \n" +
"        \"address\" : \""+address+"\", \n" +
"        \"city\" : \""+ city +"\", \n" +
"        \"state\" : \"" + state + "\", \n" +
"        \"zipcode\" : \"" + zipCode + "\", \n" +
"        \"latitude\" : \"" + latitude + "\", \n" +
"        \"longitude\" : \"" + longitude + "\", \n" +
"        \"code\" : \"" + code + "\"\n" +
"    }\n" +
"");

            
        }
    }
               
    }
    
}