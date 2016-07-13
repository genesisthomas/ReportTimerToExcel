package Salesforce.SalesforceTool;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class ExcelWriter extends ExcelDriver{
	
	public ExcelWriter(String path, String sheetName, boolean addSheet) throws Exception{
	
		super(path, sheetName, addSheet);
	
	}
	
	public void addReportKey(String sReportKey){
		try {
			lock.lock();
			this.refreshSheet();
			System.out.println("lastRow: "+this.sheet.getLastRowNum());
		XSSFRow currentRow = this.sheet.getRow(this.sheet.getLastRowNum()+1);
	
		
		XSSFCell currentCell = currentRow.createCell(currentRow.getLastCellNum());
		System.out.println("lastCell= "+currentCell.getColumnIndex());
		
		currentCell.setCellValue(sReportKey);
		System.out.println(currentCell.getStringCellValue());
		
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {lock.unlock();}		
	}
	public boolean checkIsExecutionDuplicate() throws Exception{
		lock.lock();
		boolean isDuplicate=false;
		this.refreshSheet();
		
		XSSFRow lastRow = this.sheet.getRow(this.sheet.getLastRowNum());
		XSSFRow previousRow = this.sheet.getRow(this.sheet.getLastRowNum()-1);
		
		 if(compareTwoRows(lastRow, previousRow)) {
             System.out.println("Row"+ " is duplicate");
             isDuplicate=true;
             
         } else {
             System.out.println("Row is new");
         }
		 lock.unlock();
		 return isDuplicate;
	}

	public void VerifyLastEntry() throws Exception {
		// TODO Auto-generated method stub
		if (checkIsExecutionDuplicate()) {
			System.out.print("Duplicate Row being removed");
			removeRow();
		}
			
	}
	private void RemoveDuplicateRow() throws InterruptedException {
		// TODO Auto-generated method stub
			XSSFRow lastRow = this.sheet.getRow(this.sheet.getLastRowNum());
	       this.sheet.shiftRows(this.sheet.getLastRowNum()+1, this.sheet.getLastRowNum()+1, -1);  
	       flushWorkbook();
	}
	 /**
	 * Remove a row by its index
	 * @param sheet a Excel sheet
	 * @param rowIndex a 0 based index of removing row
	 */
	public static void removeRow() {
			
	    int lastRowNum= sheet.getLastRowNum();
	    int rowIndex = lastRowNum;
	    if(rowIndex>=0&&rowIndex<lastRowNum){
	    	System.out.println("shifting rows,"+rowIndex+","+lastRowNum);
	        sheet.shiftRows(lastRowNum-1,lastRowNum, -1);
	    }
	    if(rowIndex==lastRowNum){
	    	System.out.println("shifting rows,"+rowIndex+","+lastRowNum);
	        XSSFRow removingRow=sheet.getRow(rowIndex);
	        if(removingRow!=null){
	        	System.out.println("removeRow");
	            sheet.removeRow(removingRow);
	        }
	    }
	}

	private static boolean compareTwoRows(XSSFRow row1, XSSFRow row2) {
        if((row1 == null) && (row2 == null)) {
            return true;
        } else if((row1 == null) || (row2 == null)) {
            return false;
        }
        
        int firstCell1 = row1.getFirstCellNum();
        int lastCell1 = row1.getLastCellNum();
        boolean equalRows = true;
        
        // Compare all cells in a row
        for(int i=firstCell1; i <= lastCell1; i++) {
            XSSFCell cell1 = row1.getCell(i);
            XSSFCell cell2 = row2.getCell(i);
            if(!compareTwoCells(cell1, cell2)) {
                equalRows = false;
                System.err.println("       Cell "+i+" - NOt Equal");
                break;
            } else {
                System.out.println("       Cell "+i+" - Equal");
            }
        }
        return equalRows;
    }
	
	// Compare Two Cells
    private static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2) {
        if((cell1 == null) && (cell2 == null)) {
            return true;
        } else if((cell1 == null) || (cell2 == null)) {
            return false;
        }
        
        boolean equalCells = false;
        int type1 = cell1.getCellType();
        int type2 = cell2.getCellType();
        if (type1 == type2) {
            if (cell1.getCellStyle().equals(cell2.getCellStyle())) {
                // Compare cells based on its type
                switch (cell1.getCellType()) {
                case HSSFCell.CELL_TYPE_FORMULA:
                    if (cell1.getCellFormula().equals(cell2.getCellFormula())) {
                        equalCells = true;
                    }
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    if (cell1.getNumericCellValue() == cell2
                            .getNumericCellValue()) {
                        equalCells = true;
                    }
                    break;
                case HSSFCell.CELL_TYPE_STRING:
                    if (cell1.getStringCellValue().equals(cell2
                            .getStringCellValue())) {
                        equalCells = true;
                    }
                    break;
                case HSSFCell.CELL_TYPE_BLANK:
                    if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        equalCells = true;
                    }
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    if (cell1.getBooleanCellValue() == cell2
                            .getBooleanCellValue()) {
                        equalCells = true;
                    }
                    break;
                case HSSFCell.CELL_TYPE_ERROR:
                    if (cell1.getErrorCellValue() == cell2.getErrorCellValue()) {
                        equalCells = true;
                    }
                    break;
                default:
                    if (cell1.getStringCellValue().equals(
                            cell2.getStringCellValue())) {
                        equalCells = true;
                    }
                    break;
                }
            } else {
                return false;
            }
        } else {
            return false;
        }
        return equalCells;
    }
}
