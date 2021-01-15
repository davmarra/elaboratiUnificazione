/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.provapreventivo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author david
 */
public class MyClass {
    public static void main(String[] args){
        
        String path = "C:\\Users\\david\\Desktop\\file_original\\db1111.xlsx";
        String path2 = creaElaborato(path);
        File excelFile = new File(path2);
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(excelFile);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }       
        getData(fis);       
    }
    
    
    public static String creaElaborato(String pathWB1){
        // FILE APERTURA      
        File excelFile = new File(pathWB1);
        String pathForOutput = pathWB1;
        if(!(pathForOutput.contains("_elaborato.xlsx"))){
           pathForOutput = pathWB1.replace(".xlsx", "_elaborato.xlsx");
        }
        FileInputStream fis = null;
        try {
            fis = new FileInputStream(excelFile);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(fis);
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }

        FileOutputStream output_file = null;
        try {
            output_file = new FileOutputStream(new File(pathForOutput));
        } catch (FileNotFoundException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            workbook.write(output_file); //write changes                 
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            output_file.close();  //close the stream 
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        try {
            workbook.close();
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }

        try {
            fis.close();
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
        System.out.println("Ho creato l'elaborato");
        return pathForOutput;
    }
    
    
    public static List<Object> getData(InputStream inputStream) {
        List<Object> sheetData = new ArrayList<>();
        String pathForOutput = "C:\\Users\\david\\Desktop\\file_original\\db1111_prova.xlsx";
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);            
            XSSFSheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = firstSheet.iterator();
            
            
            Row rowProva = firstSheet.getRow(0);
            
            Iterator iteratorProva = rowProva.cellIterator();
            String verifica = "Partita_IVA";
            int colVerificaInd = 0;
            while(iteratorProva.hasNext()){
                Cell cellProva = (Cell) iteratorProva.next();
                switch (cellProva.getCellType()) {
                        case STRING:
                            System.out.println("++++++ "+cellProva.getStringCellValue());
                            if(cellProva.getStringCellValue().contains(verifica)){
                                colVerificaInd = cellProva.getColumnIndex();
                                System.out.println(verifica+" TROVATO in: "+cellProva.getAddress());
                                System.out.println("Indice colonna: "+colVerificaInd);
                            }
                            break;
                        case BOOLEAN:
                            System.out.println("++++++ "+cellProva.getBooleanCellValue());
                            break;
                        case NUMERIC:                           
                            System.out.println("++++++ "+cellProva.getNumericCellValue());
                            break;
                        case BLANK:
                            System.out.println("++++++ "+"Vuoto: "+cellProva.getAddress());
                            break; 
                }
                                
            }
            
            // CODICE ELIMINAZIONE COLONNA
            /*
            int columnToDelete = 0;
            Row rowProva = firstSheet.getRow(0);
            
            Iterator iteratorProva = rowProva.cellIterator();

            while(iteratorProva.hasNext()){
                Cell cell = (Cell) iteratorProva.next();
                switch (cell.getCellType()) {
                        case STRING:
                            System.out.println("++++++ "+cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            System.out.println("++++++ "+cell.getBooleanCellValue());
                            break;
                        case NUMERIC:                           
                            System.out.println("++++++ "+cell.getNumericCellValue());
                            break;
                        case BLANK:
                            System.out.println("++++++ "+"Vuoto: "+cell.getAddress());
                            break; 
                    }
            }
            for (int rId = 0; rId <= firstSheet.getLastRowNum(); rId++) {
                Row row = firstSheet.getRow(rId);
                for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
                    Cell cOld = row.getCell(cID);
                    if (cOld != null) {
                        row.removeCell(cOld);
                    }
                Cell cNext = row.getCell(cID + 1);
                if (cNext != null) {
                    Cell cNew = row.createCell(cID, cNext.getCellType());
                    cloneCell(cNew, cNext);
                    firstSheet.setColumnWidth(cID, firstSheet.getColumnWidth(cID + 1));
                }
            }
        }
        // FINE CODICE ELIMINAZIONE COLONNA
            */
            while (iterator.hasNext()) {
                Row nextRow = iterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                               
                List<Object> row = new ArrayList<>();
                for (int colNum = 0; colNum < nextRow.getLastCellNum(); colNum++) {
                                       
                    Cell cell = nextRow.getCell(colNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    
                    switch (cell.getCellType()) {
                        case STRING:
                            row.add(cell.getStringCellValue());
                            System.out.println(cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            row.add(cell.getBooleanCellValue());
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case NUMERIC:                           
                            row.add(cell.getNumericCellValue());
                            System.out.println(cell.getNumericCellValue());
                            break;
                        case BLANK:
                            row.add("-");
                            System.out.println("Vuoto: "+cell.getAddress());
                            break; 
                    }
                }
                sheetData.add(row);                
            }
            FileOutputStream output_file = null;
        try {
            output_file = new FileOutputStream(new File(pathForOutput));
        } catch (FileNotFoundException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
        }
            workbook.write(output_file);
            output_file.close();
            workbook.close();             
            return sheetData;
        } catch (IOException ex) {
            Logger.getLogger(MyClass.class.getName()).log(Level.SEVERE, null, ex);
            return null;
        }         
    }
    
    /*
    private static void cloneCell( Cell cNew, Cell cOld ){
        cNew.setCellComment( cOld.getCellComment() );
        cNew.setCellStyle( cOld.getCellStyle() );

        switch ( cNew.getCellType() ){
            case BOOLEAN:{
                cNew.setCellValue( cOld.getBooleanCellValue() );
                break;
            }
            case NUMERIC:{
                cNew.setCellValue( cOld.getNumericCellValue() );
                break;
            }
            case STRING:{
                cNew.setCellValue( cOld.getStringCellValue() );
                break;
            }
            case ERROR:{
                cNew.setCellValue( cOld.getErrorCellValue() );
                break;
            }
            case FORMULA:{
                cNew.setCellFormula( cOld.getCellFormula() );
                break;
            }
        }
    }
    */
}
