package com.zeugor.apache_poi_excel_sample;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Set;
import java.util.Map;
import java.util.TreeMap;

public class PoiExcelFile {
	private static String inputDir = "input";
	private static String excelFile = "excelFile_demo.xlsx";
	//private static String writingFilename = "writingExcel.xlsx";
	public static void readExcel() {
		FileInputStream file = null;
		try {
			file = new FileInputStream(new File(inputDir + "/" + excelFile));

			// create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// get a sheet
			XSSFSheet sheet = workbook.getSheetAt(0);

			// iterate throught each row
			Iterator<Row> rowIterator = sheet.iterator();

			while(rowIterator.hasNext()) {
				Employ employ = new Employ();

				Row row = rowIterator.next();

				// for each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				int columnCount = 0;
				while(cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					// check the cell type
					switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							int intCell = Math.round((float) cell.getNumericCellValue());
							if (columnCount == 0) {
								employ.setId(intCell);
							}
							if (columnCount == 2) {
								employ.setSalary(intCell);
							}
							break;
						case Cell.CELL_TYPE_STRING:
							employ.setName(cell.getStringCellValue());
							break;
					}
					columnCount++;
				}

				System.out.println(employ);
			}
			workbook.close();
            System.out.println("readingExcel.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (file != null) {
					file.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public static void writeExcel() {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employee Data");

        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "SALARY"});
        data.put("2", new Object[] {1, "Amit", 1000});
        data.put("3", new Object[] {2, "Lokesh", 2000});
        data.put("4", new Object[] {3, "John", 3000});
        data.put("5", new Object[] {4, "Brian", 4000});	
        
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset) {
        	Row row = sheet.createRow(rownum++);
        	Object[] objArr = data.get(key);
        	int cellnum = 0;
        	for (Object obj : objArr) {
        		Cell cell = row.createCell(cellnum++);
        		if (obj instanceof String) {
        			cell.setCellValue((String)obj);
        		} else if (obj instanceof Integer) {
        			cell.setCellValue((Integer)obj);
        		}
        	}
        }

		FileOutputStream out = null;
		try { 
			out = new FileOutputStream(new File(inputDir + "/" + excelFile));
			workbook.write(out);
			workbook.close();
            System.out.println("writingExcel.xlsx written successfully on disk.");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null) {
					out.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}			
	}

	public static void writingFormulasExcel() {
	    XSSFWorkbook workbook = new XSSFWorkbook();
	    XSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");
	  
	    Row header = sheet.createRow(0);
	    header.createCell(0).setCellValue("Pricipal");
	    header.createCell(1).setCellValue("RoI");
	    header.createCell(2).setCellValue("T");
	    header.createCell(3).setCellValue("Interest (P r t)");
	      
	    Row dataRow = sheet.createRow(1);
	    dataRow.createCell(0).setCellValue(14500d);
	    dataRow.createCell(1).setCellValue(9.25);
	    dataRow.createCell(2).setCellValue(3d);
	    dataRow.createCell(3).setCellFormula("A2*B2*C2");
	      
	    try {
	        FileOutputStream out =  new FileOutputStream(new File(inputDir + "/formulaDemo.xlsx"));
	        workbook.write(out);
	        out.close();
	        System.out.println("Excel with foumula cells written successfully");
	          
	    } catch (Exception e) {
	        e.printStackTrace();
		}
	}

	public static void readingFormulasExcel() {
		try {
	        FileInputStream file = new FileInputStream(new File(inputDir + "/formulaDemo.xlsx"));
	 
	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);
	 
	        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
	         
	        //Get first/desired sheet from the workbook
	        XSSFSheet sheet = workbook.getSheetAt(0);
	 
	        //Iterate through each rows one by one
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();
	             
	            while (cellIterator.hasNext())
	            {
	                Cell cell = cellIterator.next();
	                //Check the cell type after eveluating formulae
	                //If it is formula cell, it will be evaluated otherwise no change will happen
	                switch (evaluator.evaluateInCell(cell).getCellType())
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(cell.getStringCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_FORMULA:
	                        //Not again
	                        break;
	                }
	            }
	            System.out.println("");
	        }
	        file.close();
	    }
	    catch (Exception e)
	    {
	        e.printStackTrace();
	    }		
	}

	public static void main(String... arg) throws IOException {
		System.out.println("main...");
		writeExcel();
		readExcel();
		writingFormulasExcel();		
		readingFormulasExcel();		
	}
}
