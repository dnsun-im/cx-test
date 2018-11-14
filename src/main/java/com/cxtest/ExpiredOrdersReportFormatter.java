package com.cxtest;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.log4j.Logger;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
/**
 * Takes a CareEvolve expired orders report and filters the records to only include those for the next month.
 * Run this from command line. Takes parameters in this order:
 * 
 * 0) Run date in DD-MM-YYYY format. Includes all the records for the month after the run date.
 * 1) the path to the report
 * 2) the target directory for the formatted spreadsheet
 * 
 * The result file name will be named ExpiredOrders_MONTH_YEAR.xls.
 * 
 * Example to call:
 * java com.caredx.utilities.careevolve.ExpiredOrdersReportFormatter 18-11-2018 c:\temp\CareEvolveReport.xls c:\customercare\reports
 */
public class ExpiredOrdersReportFormatter {
	private static Logger logger = Logger.getLogger(ExpiredOrdersReportFormatter.class);
	public static final int DOB_INDEX = 2;
	public static final int EXPIRE_DATE_INDEX = 5;
	public static final int COLUMN_COUNT = 8;
	
	private static final DateTimeFormatter EURO_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");
	private static final DateTimeFormatter EXPIRED_FORMAT = DateTimeFormatter.ofPattern("MMM d yyyy h:mma");
	private static final DateTimeFormatter TARGET_FORMAT = DateTimeFormatter.ofPattern("MMM dd, yyyy");
 
	private LocalDateTime start;
	private LocalDateTime end;
	private String reportPath;
	private String targetDirectory;
	private Sheet sheet;
	private Workbook workbook; 
	
	public static void main(String[] args) { 
		try {
			LocalDate date = LocalDate.parse(args[0], EURO_FORMAT);
			String reportPath = args[1];
			String targetDirectory = args[2]; 
			logger.info("Formatting " + reportPath + " " + " for run date: " + date.format(TARGET_FORMAT) );
			
			ExpiredOrdersReportFormatter formatter = new ExpiredOrdersReportFormatter(date, reportPath, targetDirectory);
			formatter.format();
			logger.info("Formatting complete.");
		}
		catch(Exception e) {
			logger.error("Error running the expired orders formatter.", e);
		}
		
		System.exit(0);
	}
	
	public ExpiredOrdersReportFormatter(LocalDate date, String reportPath, String targetDirectory) throws FileNotFoundException, BiffException, IOException {
		this.reportPath = reportPath;
		this.targetDirectory = targetDirectory;
		this.start = date.atStartOfDay().with(TemporalAdjusters.firstDayOfNextMonth());
		this.end =  start.with(TemporalAdjusters.lastDayOfMonth()).plusDays(1l);
		
	    workbook = getWorkbook();
		sheet = workbook.getSheet(0);
	}
	
	public void format() throws FileNotFoundException, IOException, RowsExceededException, WriteException {
		String newXlsPath = targetDirectory + "\\ExpiredOrders_" + start.getMonth() + "_" + start.getYear()  + ".xls";
		WritableWorkbook newXls = Workbook.createWorkbook(new File(newXlsPath));
		WritableSheet newSheet = newXls.createSheet("Expired Orders", 0);
		initColumns(newSheet);
		Map<String, List<List<Cell>>> filtered = getFiltered(); 
		
		int r = 1;
		for(Entry<String, List<List<Cell>>> entry : filtered.entrySet()) {
			if(entry.getValue().size() > 0) {
				// Add the practice row
				Label practiceLabel = new Label(0, r, entry.getKey(), getCellFormat(r));
				newSheet.addCell(practiceLabel);
				for (int b=1; b < COLUMN_COUNT; b++) {
					Label cellLabel = new Label(b, r, "", getCellFormat(r));
					newSheet.addCell(cellLabel);
				}
				r++;
				
				// Add patients
				for(List<Cell> row : entry.getValue()) {
					for(Cell cell : row) {
						int col = cell.getColumn();
						String cellValue = DOB_INDEX == col ? formatDob(cell): cell.getContents() ;   
						Label cellLabel = new Label(col, r, cellValue, getCellFormat(r));
						newSheet.addCell(cellLabel);
					}
					r++;
				}
			}
		}
		
		newXls.write();
		newXls.close();
		logger.info("Formatted Excel file saved to " + newXlsPath);
	}
	
	private Map<String, List<List<Cell>>> getFiltered() {
		//reads the excel, group by practice including only patients in the date range
		Map<String, List<List<Cell>>> output = new LinkedHashMap<>();
		
		int rcount = sheet.getRows();
		String currentPractice = "";
		
		for(int r=2; r < rcount; r++) {
			String practice = sheet.getCell(0,r).getContents(); 
			if(practice.length() > 0) {
				if(!output.containsKey(practice)) {
					List<List<Cell>> rows = new ArrayList<>();
					output.put(practice, rows);
				}
				currentPractice = practice;
				continue;
			} 
			
			if(shouldKeep(r)) {
				List<Cell> row = new ArrayList<>();
				for (int c=0; c < COLUMN_COUNT; c++) {
					Cell cell = sheet.getCell(c, r);
					row.add(cell); 
				}
				output.get(currentPractice).add(row);
			}
		}
		return output;
	}
	
	
	private void initColumns(WritableSheet writable) throws WriteException {
		for(int w=0; w<COLUMN_COUNT; w++) {
			String colHeading = sheet.getCell(w, 1).getContents();
			Label colLabel = new Label(w, 0, colHeading, getCellFormat(0));
			writable.addCell(colLabel);
			writable.setColumnView(w, 25);
		}
	}
	
	private WritableCellFormat getCellFormat(int row) throws WriteException {
		WritableCellFormat format = new WritableCellFormat();
        WritableFont font = (0==row) ? new WritableFont(WritableFont.ARIAL, 11, WritableFont.BOLD)
        						   	 : new WritableFont(WritableFont.ARIAL, 10);
        format.setFont(font);
        format.setBorder(Border.ALL, BorderLineStyle.THIN);
        return format; 
	} 
	
	private boolean shouldKeep(int row) {
		Cell expirationCell = sheet.getCell(EXPIRE_DATE_INDEX, row);
		boolean should = true;
		if(expirationCell == null || expirationCell.getContents().equals("")) { 
			should = true;
		}
		else {
			try {
				String expText = expirationCell.getContents().replace("  ", " "); 
				LocalDateTime expiredDate = LocalDateTime.parse(expText, EXPIRED_FORMAT);
				should =  expiredDate.isAfter(start) && expiredDate.isBefore(end);
			}
			catch (Exception e) {
				logger.debug("Ignoring date parsing errors", e);
			}
		} 
		return should;
	}
	
	private Workbook getWorkbook() throws FileNotFoundException, IOException, BiffException {
		Workbook workbook =  Workbook.getWorkbook(new File(reportPath)); 
		return workbook;
	}
	
	private String formatDob(Cell cell) {
		String formatted = "";
		try {
			String dobCell = cell.getContents().trim();
			if(dobCell==null || "".equals(dobCell)) {
				return formatted;
			}
			formatted = LocalDate.parse(dobCell, EURO_FORMAT).format(TARGET_FORMAT); 
		}
		catch(DateTimeParseException e) {
			// ignore return empty space
			logger.debug("Ignore this DOB for this cell ", e);
		} 
		return formatted;
	} 
}
