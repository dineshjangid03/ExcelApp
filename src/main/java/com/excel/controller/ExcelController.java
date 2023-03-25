package com.excel.controller;

import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import com.excel.exception.QueryNotFoundException;
import com.excel.service.ExcelService;

import java.io.IOException;

@Controller
public class ExcelController {
	
	@Autowired
	private ExcelService excelService;
	
	
	@GetMapping("/search")
	public String search(Model model) {
	    return "search";
	}
	
    @PostMapping("/search")
    public String search(@RequestParam("query") String query, Model model) throws IOException {
        
		try {
			Sheet resultSheet = excelService.search(query);
			model.addAttribute("results", resultSheet);
		} catch (QueryNotFoundException e) {
			model.addAttribute("error", "No results found.");
		}
        
        return "searchResults";
    
    }
}
