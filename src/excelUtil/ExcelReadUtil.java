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
 * ��ȡExcel�ļ�  (!!!����С�ļ�!!!)
 */
public class ExcelReadUtil {

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	/**
	 *�ļ���׺��(xls,xlsx) 
	 */
	protected String extension = "";
	
	/**
	 * �ļ���
	 */
	protected String fileName = "";
	
	protected String filePath ="";

	/**
	 * Excel�ļ�
	 */
	protected File excelFile;
	
	/**
	 * ����Excel��Workbook����
	 */
	protected Workbook workbook = null;
	
	/**
	 * ���ڶ�ȡ��ʽ
	 */
	protected FormulaEvaluator evaluator = null;
	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	
	/**
	 * ������Excel�ļ��Ĺ���
	 * @param excelPath  Excel�ļ���ȡ·��
	 * @throws ExcelNotExcelFileException �ļ�����Excel�ļ�
	 * @throws ExcelNullParameterException ����Ϊnull
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
			//���ַ�ʽ Excel 2003/2007/2010 ���ǿ��Դ����  
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
	 * ��ȡExcel�ļ���ʽ(.xls��.xlsx)
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
	 * ��ȡ�ļ���sheet�������
	 * @return
	 */
	public int getSheetSize(){
		return this.workbook.getNumberOfSheets();
	}
	
	/**
	 * ��ȡ��һ��sheet
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
	 * ��ȡ�±�ΪIndex��Sheet
	 * @param index sheet���±�ֵ
	 * @throws ExcelIndexOutOfBoundsException ����С��0�����sheet��������
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
	 * ͨ�����ƶ�ȡsheet
	 * @param name sheet������
	 * @throws ExcelIllegalArgumentException ָ������sheet������
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
	 * �رն�ȡ����
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
	 * ����ļ��Ƿ�ΪExcel�ļ�
	 * @param filePath �ļ�·��
	 * @return ����ļ�Ϊxls,xlsx��ʽ�򷵻�true,����false
	 * @throws ExcelNullParameterException �ļ�·��Ϊnull
	 */
	public static boolean isExcelFile(String filePath) throws ExcelNullParameterException{
		if(filePath == null){
			throw new ExcelNullParameterException();
		}
		String ext = filePath.substring(filePath.lastIndexOf("."));
		return ext.equals(".xlsx") || ext.equals(".xls") ? true : false;
	}
	
	/**
	 * ͨ���ļ�·�������ļ���(����׺)
	 * @param filePath
	 * @return
	 */
	public static String getFileName(String filePath){
		return filePath.substring(filePath.lastIndexOf("\\")+1, filePath.length());
	}
	
}
	
	