import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) throws Exception {
      
    
    //declaration de variable de type file et specefier le chemin
    File myFile = new File("C:/Users/PC/Desktop/CCNB/2eme/PROG1284/JavaExcelFile/staff.xlsx");
                FileInputStream fis = new FileInputStream(myFile);

                // creer un istence XSSFWorkbook pour le pouvoir manipuler das le program java
                XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
            
                // creer un instence XSSFsheet pour retourner la feuille excel 
                XSSFSheet mySheet = myWorkBook.getSheetAt(0);
            
                // variable de maipulation des lignes de fichier excel
                Iterator<Row> rowIterator = mySheet.iterator();


                //saisir les ligne a la fin de fichier excel
                Map<String, Object[]> data = new HashMap<String, Object[]>();
                data.put("7", new Object[] {7d, "Sonya", "75K", "SALES", "Rupert"});
                data.put("8", new Object[] {8d, "Kris", "85K", "SALES", "Rupert"});
                data.put("9", new Object[] {9d, "Dave", "90K", "SALES", "tankeu"});
            
                // saisir le variale qui va avoir les lignes de fichier
                Set<String> newRows = data.keySet();
            
                // lire lenombre de derier ligne dans le fichier         
                int rownum = mySheet.getLastRowNum();         
            
                for (String key : newRows) {
                
                    // creation de la ligne dans le variable de type mysheet
                    Row row = mySheet.createRow(rownum++);
                    Object [] objArr = data.get(key);
                    int cellnum = 0;
                    for (Object obj : objArr) {
                        Cell cell = row.createCell(cellnum++);
                        if (obj instanceof String) {
                            cell.setCellValue((String) obj);
                        } else if (obj instanceof Boolean) {
                            cell.setCellValue((Boolean) obj);
                        } else if (obj instanceof Date) {
                            cell.setCellValue((Date) obj);
                        } else if (obj instanceof Double) {
                            cell.setCellValue((Double) obj);
                        }
                    }
                }
            
                // affecter les valeur das le fichier excel dans le fichier das la machine
                FileOutputStream os = new FileOutputStream(myFile);
                myWorkBook.write(os);
                System.out.println("Writing on XLSX file Finished ...");
            }
        }

    

