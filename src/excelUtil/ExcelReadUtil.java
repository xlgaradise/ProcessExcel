package excelUtil;
/** 
 * @author HPC
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

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
	
	/**
	 * 最近一次读取的sheet列表
	 */
	protected ArrayList<Sheet> sheetList = null; 	
	
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
	 * 获取读取的sheet列表
	 * @return sheet列表
	 */
	public ArrayList<Sheet> getSheetList() {
		return sheetList;
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
	
	/**
	 * 读取第一个sheet
	 */
	public void readFirstSheet(){
		try {
			this.sheetList = getSheetList(0, 1);
		} catch (ExcelIndexOutOfBoundsException e) {
		}
	}
	
	/**
	 * 读取下标为Index的Sheet
	 * @param index sheet的下标值
	 * @throws ExcelIndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetByIndex(int index) throws ExcelIndexOutOfBoundsException{
		this.sheetList = getSheetList(index, 1);
	}
	
	/**
	 * 通过名称读取sheet
	 * @param name sheet的名称
	 * @throws ExcelIllegalArgumentException 名称错误，无法获取指定sheet
	 */
	public void readSheetByName(String name) throws ExcelIllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			sheetList.add(sheet);
		}
		else
			throw new ExcelIllegalArgumentException();
	}
	
	/**
	 * 读取指定范围的sheet列表
	 * @param startIndex sheet开始的下标值
	 * @param length 要读取sheets的长度
	 * @throws ExcelIndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetList(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		this.sheetList = getSheetList(startIndex, length);
	}
	
	/**
	 * 读取所有的sheet
	 */
	public void readAllSheet(){
		int sheetCount = workbook.getNumberOfSheets();
        Sheet sheet = null;
        for(int i=0;i<sheetCount;i++){
        	sheet = workbook.getSheetAt(i);
        	sheetList.add(sheet);
        }
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
	 * 获取需要的sheet列表
	 * @param startIndex sheet开始的下标值
	 * @param length 要读取sheets的长度,如果长度过长则读取至数据结尾
	 * @return Sheet 列表
	 * @throws ExcelIndexOutOfBoundsException 参数错误
	 */
	protected ArrayList<Sheet> getSheetList(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量  
        int endIndex = 0; 
        try {
			endIndex = isIndexOutOfBounds(sheetCount, startIndex, length);
		} catch (ExcelIndexOutOfBoundsException e) {
			throw e;
		}
        for(int i=startIndex;i<=endIndex;i++){
        	sheetList.add(workbook.getSheetAt(i));
        }
        return sheetList;
	}
	
	/**
	 * 判断数据长度、起始下标和读取长度参数是否越界
	 * @param count 数据总长度,不能小于1
	 * @param startIndex 起始下标不能小于零或大于最大值
	 * @param length 读取的长度,不能小于0
	 * @return 如果参数越界抛出异常,否则返回要读取的最后一个下标值(如果读取长度大于总长度,则返回最大下标值)
	 * @throws ExcelIndexOutOfBoundsException 参数越界错误
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		if(count<1){ //数据长度小于1
			throw new ExcelIndexOutOfBoundsException();
		}
		if(length<0){//读取长度小于零
			throw new ExcelIndexOutOfBoundsException();
        }
        if(startIndex > count -1 || startIndex < 0){//开始下标大于最大的下标值或小于零
        	throw new ExcelIndexOutOfBoundsException();
        }
        //要读取的最后一个下标,如果下标越界，则读取至最后一个值
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
}
	
	