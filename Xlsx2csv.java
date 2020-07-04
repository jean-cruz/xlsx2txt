/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlsx2csv;

/**
 *
 * @author jean.cruz
 */
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Xlsx2csv {

     private static void convertSelectedSheetInXLXSFileToCSV(File xlsxFile, int sheetIdx) throws Exception {
         
            FileInputStream fileInStream = new FileInputStream(xlsxFile);
     
            //Open the xlsx and get the requested sheet from the workbook8
            XSSFWorkbook workBook = new XSSFWorkbook(fileInStream);
            XSSFSheet selSheet = workBook.getSheetAt(sheetIdx);
     
            // Iterate through all the rows in the selected sheet
            Iterator<Row> rowIterator = selSheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next(); 
                // Iterate through all the columns in the row and build ","
                // separated string
                Iterator<Cell> cellIterator = row.cellIterator();
                StringBuffer sb = new StringBuffer();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    if (sb.length() != 0) {
                        sb.append(",");
                    }
                     
                    // If you are using poi 4.0 or over, change it to
                    // cell.getCellType
                    switch (cell.getCellTypeEnum()) {
                    case BLANK:
                        sb.append("BLANK");
                        break;        
                    case STRING:
                        sb.append(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        sb.append(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        sb.append(cell.getBooleanCellValue());
                        break;
                    
                    default:
                       sb.append("default");                       
                    }
                }
           
                System.out.println(sb.toString());
            }
            workBook.close();
        }
        /**
        * Processa a linha de comando.
        * Parâmetros esperados:
        *      -modelo LH | STKS
        *      -inputTXT {arquivo}
        * @param args the command line arguments
     * @throws java.lang.Exception
        */
        public static void main(String[] args) throws Exception{
            try{
                if(args.length == 0){
                    exibeUso();
                    System.exit(1);
                    return;
                }
                // Recupera os parâmetros da linha de comando
                HashMap<String, String> cmdArgs = new HashMap<String, String>();
                for(int i = 0; i < args.length; i += 2){
                    cmdArgs.put(args[i], args[i + 1]);
                }
                // Recupera a localização do arquivo de dados
                if(!cmdArgs.containsKey("-inputTXT")){
                    throw new Exception("Arquivo de entrada não informado.");
                }
                String nomeArquivo = cmdArgs.get("-inputTXT");
                if(!(new File(nomeArquivo)).exists()){
                    throw new Exception("Arquivo de dados não encontrado.");
                }
                File myFile = new File(nomeArquivo); //"C:\\temp\\teste\\teste.xlsx"
                int sheetIdx = 0; // 0 for first sheet
                convertSelectedSheetInXLXSFileToCSV(myFile, sheetIdx);
            }catch(Exception e) { // Exceções geradas internamente
                // Envia para a saída padrão a pilha de erros de execução.
                e.printStackTrace(System.out);
                // Sinaliza para o SO que o programa não finalizou com sucesso.
                System.exit(-1);
            
            }
        }     
        private static void exibeUso() {
            System.out.println("Parâmetros esperados:");
            System.out.println("    -inputTXT [arquivo]");
            System.out.println("");
        }
}