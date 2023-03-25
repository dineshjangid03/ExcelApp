package com.excel.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Controller
public class ExcelController {
	
	
	@GetMapping("/search")
	public String search(Model model) {
	    return "search";
	}
	
    @PostMapping("/search")
    public String search(@RequestParam("query") String query, Model model) throws IOException {
        
    	// address of file path
        String filePath="src/main/resources/static/sample.xlsx";
        FileInputStream inputStream=new FileInputStream(filePath);
        Workbook workbook=new XSSFWorkbook(inputStream);
        //getting sheet from file
        Sheet sheet=workbook.getSheetAt(0);
        
        // getting result sheet from excel file
        Sheet resultSheet=workbook.getSheet("Result");
        //checking that sheet is available or it's null
        if (resultSheet==null) {
        	//if it's null we have to create new sheet to store result
            resultSheet=workbook.createSheet("Result");
        }
        
        //taking flag for search query is present or not, initially it's false if we found we make it as true
        boolean flag=false;

        //Iterating through the rows of the sheet 
        for (Row row : sheet) {
        	//Iterating through every cell of rows
            for (Cell cell : row) {
            	// there is many type of cell like String Numeric Boolean
            	// if cell is String type that mean it contains some query, so we will further compare with our input query
                if (cell.getCellType()==CellType.STRING) {
                	// taking value of cell
                	String cellValue=cell.getStringCellValue();
                	// comparing every cells value with input query
                	if (cellValue.equals(query)) {// if we found same query this block will execute
                		//here i'm going to create new row in result sheet to store our entire row of query
                        Row newRow=resultSheet.createRow(resultSheet.getLastRowNum() + 1);
                        // iterating throw row which contains query
                        for (Cell originalCell : row) {
                        	// creating new cell in newRow to store values
                            Cell newCell=newRow.createCell(originalCell.getColumnIndex());
                            //if value is type of string I'm setting value in newCell using getStringCellValue method
                            if(originalCell.getCellType()==CellType.STRING)
                            	newCell.setCellValue(originalCell.getStringCellValue());
                            //if value is type of Numeric I'm setting value in newCell using getNumericCellValue method
                            else if(originalCell.getCellType()==CellType.NUMERIC)
                            	newCell.setCellValue(originalCell.getNumericCellValue());
                            // we can also check for getBooleanCellValue or getDateCellValue ...
                        }
                        //making flag as true so further it indicate that we found query in file
                        flag=true;
                    }
                }
            }
        }
        
        
        if(flag)
        	model.addAttribute("results", resultSheet);
        else 
        	model.addAttribute("error", "No results found.");
        
        
        FileOutputStream outputStream=new FileOutputStream(filePath);
        workbook.write(outputStream);
        
        outputStream.close();
        workbook.close();
        
        
        return "searchResults";
    }
}
