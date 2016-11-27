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

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 读取Excel文件,仅可获取Excel文件下的各个sheet表
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
	
	protected FormulaEvaluator evaluator = null;
	
	/**
	 * 最近一次读取的sheet列表
	 */
	protected ArrayList<Sheet> sheetList = null; 	
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	

	/**
	 * 创建读Excel文件的工具
	 * @param excelPath  Excel文件读取路径
	 * @throws IllegalArgumentException 文件不存在或格式错误
	 * @throws NullPointerException 文件路径为null
	 * @throws SecurityException 文件拒绝访问
	 * @throws FileNotFoundException 文件读取出错
	 * @throws Exception 生成工作薄出错
	 */
	public ExcelReadUtil(String excelPath) throws IllegalArgumentException,NullPointerException,
						SecurityException,FileNotFoundException,Exception{
		try {
			if(isExcelFile(excelPath)){
				this.excelFile = new File(excelPath);
				this.filePath = excelPath;
				String name = this.excelFile.getName();
				this.extension = name.substring(name.lastIndexOf("."));
				this.fileName = name;
				//这种方式 Excel 2003/2007/2010 都是可以处理的  
				if(extension.equals(".xls")){
					FileInputStream is = new FileInputStream(excelFile); 
					this.workbook = new HSSFWorkbook(is);
					is.close();
				}else{//.xlsx
					FileInputStream is = new FileInputStream(excelFile); 
					this.workbook = new XSSFWorkbook(is);
					is.close();
				}
		        evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
			}else {
				throw new IllegalArgumentException("文件不是Excel文件");
			}
		}catch (IllegalArgumentException e) {
			throw e;
		}catch (NullPointerException e) {
			throw e;
		}catch (SecurityException e) {
			throw e;
		} catch (FileNotFoundException e) {
			throw e;
		} catch (EncryptedDocumentException e) {
			throw e;
		} catch (IOException e) {
			throw e;
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
	 * @return 如果文件为Excel格式则返回true,否则false
	 * @throws IllegalArgumentException 文件不存在
	 * @throws NullPointerException 文件路径为null
	 * @throws SecurityException 文件拒绝访问
	 */
	public static boolean isExcelFile(String filePath) throws IllegalArgumentException,
															NullPointerException,SecurityException{
		try {
			File file = new File(filePath);
			if(!file.exists()){
				throw new IllegalArgumentException("路径错误,文件不存在");
			}else{
				String name = file.getName();
				String ext = name.substring(name.lastIndexOf("."));
				if(ext.equals(".xls") || ext.equals(".xlsx")) return true;
				else return false;
			}
		} catch (NullPointerException e) {
			throw new NullPointerException("文件路径为空");
		} catch (SecurityException  e) {	
			throw new SecurityException("文件拒绝访问");
		}
	}
	
	/**
	 * 读取第一个sheet
	 */
	public void readFirstSheet(){
		this.sheetList = getSheetList(0, 1);
	}
	
	/**
	 * 读取下标为Index的Sheet
	 * @param index sheet的下标值
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetByIndex(int index) throws IndexOutOfBoundsException{
		try {
			this.sheetList = getSheetList(index, 1);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 通过名称读取sheet
	 * @param name sheet的名称
	 * @throws IllegalArgumentException 名称错误，无法获取指定sheet
	 */
	public void readSheetByName(String name) throws IllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			sheetList.add(sheet);
		}
		else
			throw new IllegalArgumentException("无法获取指定名称的sheet");
	}
	
	/**
	 * 读取指定范围的sheet列表
	 * @param startIndex sheet开始的下标值
	 * @param length 要读取sheets的长度
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		try {
			this.sheetList = getSheetList(startIndex, length);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
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
	 * @throws IndexOutOfBoundsException 参数错误
	 */
	protected ArrayList<Sheet> getSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量  
        try {
			int endIndex = isIndexOutOfBounds(sheetCount, startIndex, length);
			for(int i=startIndex;i<=endIndex;i++){
	        	sheetList.add(workbook.getSheetAt(i));
	        }
	        return sheetList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 判断数据长度、起始下标和读取长度参数是否越界
	 * @param count 数据总长度,不能小于1
	 * @param startIndex 起始下标不能小于零或大于最大值
	 * @param length 读取的长度,不能小于0
	 * @return 如果参数越界抛出异常,否则返回要读取的最后一个下标值(如果下标越界,则返回最大下标值)
	 * @throws IndexOutOfBoundsException 参数错误
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("数据长度小于1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("读取长度小于零");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("开始下标大于最大的下标值或小于零");
        }
        //要读取的最后一个下标,如果下标越界，则读取至最后一个值
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
}
	
	