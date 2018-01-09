package excelUtil;
/** 
 * @author HPC
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNotExcelFileException;
import exception.ExcelNullParameterException;


/**
 * 读取Excel文件  (!!!仅限小文件!!!)
 */
public class ExcelReadUtil {

	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	/**
	 *文件后缀名(xls,xlsx) 
	 */
	protected String extension = "";
	
	/**
	 * 文件名
	 */
	protected String fileName = "";
	
	protected String filePath ="";

	/**
	 * Excel文件
	 */
	protected File excelFile;
	
	/**
	 * 操作Excel的Workbook工具
	 */
	protected Workbook workbook = null;
	
	/**
	 * 用于读取公式
	 */
	protected FormulaEvaluator evaluator = null;
	
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	

	
	/**
	 * 创建读Excel文件的工具
	 * @param excelPath  Excel文件读取路径
	 * @throws ExcelNotExcelFileException 文件不是Excel文件
	 * @throws ExcelNullParameterException 参数为null
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public ExcelReadUtil(String excelPath) throws ExcelNotExcelFileException,
											ExcelNullParameterException,FileNotFoundException,IOException{
		
		boolean isExcel = isExcelFile(excelPath);
		if(isExcel){
			this.excelFile = new File(excelPath);
			this.filePath = excelPath;
			String name = this.excelFile.getName();
			this.extension = name.substring(name.lastIndexOf("."));
			this.fileName = name;
			//这种方式 Excel 2003/2007/2010 都是可以处理的  
			FileInputStream is = null;
			if (extension.equals(".xls")) {
				is = new FileInputStream(excelFile);
				this.workbook = new HSSFWorkbook(is);
			} else {// .xlsx
				is = new FileInputStream(excelFile);
				this.workbook = new XSSFWorkbook(is);
			}
			if (is != null)
				is.close();
	        evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
		}else{
			throw new ExcelNotExcelFileException();
		}
	}
	
	/**
	 * 获取Excel文件格式(.xls或.xlsx)
	 * @return
	 */
	public String getExtension() {
		return extension;
	}
	
	public String getFileName(){
		return fileName;
	}
	
	public String getFilePath(){
		return filePath;
	}
	
	public Workbook getWorkBook(){
		return workbook;
	}
	
	public FormulaEvaluator getEvaluator(){
		return evaluator;
	}
	
	/**
	 * 获取文件中sheet表的数量
	 * @return
	 */
	public int getSheetSize(){
		return this.workbook.getNumberOfSheets();
	}
	
	/**
	 * 读取第一个sheet
	 */
	public SheetReadUtil readFirstSheet(){
		try {
			Sheet sheet = workbook.getSheetAt(0);
			return new SheetReadUtil(sheet);
		} catch (ExcelNullParameterException e) {
			return null;
		}
	}
	
	/**
	 * 读取下标为Index的Sheet
	 * @param index sheet的下标值
	 * @throws ExcelIndexOutOfBoundsException 参数小于0或大于sheet的总数量
	 */
	public SheetReadUtil readSheetByIndex(int index) throws ExcelIndexOutOfBoundsException{
        try {
        	Sheet sheet = workbook.getSheetAt(index);
			return new SheetReadUtil(sheet);
		} catch (IllegalArgumentException | ExcelNullParameterException e) {
			throw new ExcelIndexOutOfBoundsException();
		}
	}
	
	/**
	 * 通过名称读取sheet
	 * @param name sheet的名称
	 * @throws ExcelIllegalArgumentException 指定名称sheet不存在
	 */
	public SheetReadUtil readSheetByName(String name) throws ExcelIllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			try {
				return new SheetReadUtil(sheet);
			} catch (ExcelNullParameterException e) {
				return null;
			}
		}
		else
			throw new ExcelIllegalArgumentException();
	}
	
	/**
	 * 关闭读取工具
	 * @throws IOException
	 */
	public void close() throws IOException{
		try {
			this.workbook.close();
		} catch (IOException e) {
			throw e;
		}
	}
	
	/**
	 * 检查文件是否为Excel文件
	 * @param filePath 文件路径
	 * @return 如果文件为xls,xlsx格式则返回true,否则false
	 * @throws ExcelNullParameterException 文件路径为null
	 */
	public static boolean isExcelFile(String filePath) throws ExcelNullParameterException{
		if(filePath == null){
			throw new ExcelNullParameterException();
		}
		String ext = filePath.substring(filePath.lastIndexOf("."));
		return ext.equals(".xlsx") || ext.equals(".xls") ? true : false;
	}
	
	/**
	 * 通过文件路径返回文件名(带后缀)
	 * @param filePath
	 * @return
	 */
	public static String getFileName(String filePath){
		return filePath.substring(filePath.lastIndexOf("\\")+1, filePath.length());
	}
	
}
	
	