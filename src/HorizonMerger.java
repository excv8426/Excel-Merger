import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.UUID;

import javax.activation.UnsupportedDataTypeException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class HorizonMerger {
	
	private String inputpath;
	private String outputpath;
	
	public HorizonMerger(String inputpath,String outputpath){
		
		
		this.inputpath=inputpath;
		this.outputpath=outputpath+"\\"+UUID.randomUUID()+".xlsx";
	}
	
	public void merge() throws IOException{
		File dir=new File(inputpath);
		File[] excels=dir.listFiles(new ExcelFilefilter());
		FileInputStream excelStream=null;
		Workbook inputWorkbook=null;
		SXSSFWorkbook outputWorkbook=new SXSSFWorkbook(100);
		Sheet inputSheet=null;
		Sheet outputSheet=outputWorkbook.createSheet();
		int rowcount=0;
		for (File file : excels) {
			
			excelStream=new FileInputStream(file);
			inputWorkbook=createWorkbook(excelStream);
			inputSheet=inputWorkbook.getSheetAt(0);
			for (int i = 0; i <= inputSheet.getLastRowNum(); i++) {
				
				if ((rowcount>0)&&(i==0)) {
					continue;
				}
				rowCopy(inputSheet.getRow(i), outputSheet.createRow(rowcount));
				rowcount++;	
			}
		}
		FileOutputStream outputStream=new FileOutputStream(new File(outputpath));
		outputWorkbook.write(outputStream);
		outputWorkbook.close();
		outputStream.flush();
		outputStream.close();
	}
	
	private void rowCopy(Row rsc,Row dst) throws UnsupportedDataTypeException{
		Iterator<Cell> iterator=rsc.cellIterator();
		int i=0;
		Cell rscCell;
		Cell dstCell;
		while (iterator.hasNext()) {
			rscCell=iterator.next();
			dstCell=dst.createCell(i);
			i++;
			switch (rscCell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				dstCell.setCellType(Cell.CELL_TYPE_STRING);
				dstCell.setCellValue(rscCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				dstCell.setCellType(Cell.CELL_TYPE_STRING);
				dstCell.setCellValue(String.valueOf(rscCell.getNumericCellValue()));
				break;
			default:
				System.out.println(rscCell.getCellType());
				throw new UnsupportedDataTypeException("未知的Cell类型。");
			}
			
		}
	}
	
	private Workbook createWorkbook(InputStream inputStream){
		Workbook wb=null;
		try {
			wb = WorkbookFactory.create(inputStream);
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} 
		return wb;
	}
}
