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
	
	/**
	 * ���һ�ζ�ȡ��sheet�б�
	 */
	protected ArrayList<Sheet> sheetList = null; 	
	
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
	 * ��ȡ��ȡ��sheet�б�
	 * @return sheet�б�
	 */
	public ArrayList<Sheet> getSheetList() {
		return sheetList;
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
	
	/**
	 * ��ȡ��һ��sheet
	 */
	public void readFirstSheet(){
		try {
			this.sheetList = getSheetList(0, 1);
		} catch (ExcelIndexOutOfBoundsException e) {
		}
	}
	
	/**
	 * ��ȡ�±�ΪIndex��Sheet
	 * @param index sheet���±�ֵ
	 * @throws ExcelIndexOutOfBoundsException ����Խ�����
	 */
	public void readSheetByIndex(int index) throws ExcelIndexOutOfBoundsException{
		this.sheetList = getSheetList(index, 1);
	}
	
	/**
	 * ͨ�����ƶ�ȡsheet
	 * @param name sheet������
	 * @throws ExcelIllegalArgumentException ���ƴ����޷���ȡָ��sheet
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
	 * ��ȡָ����Χ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���
	 * @throws ExcelIndexOutOfBoundsException ����Խ�����
	 */
	public void readSheetList(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		this.sheetList = getSheetList(startIndex, length);
	}
	
	/**
	 * ��ȡ���е�sheet
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
	 * ��ȡ��Ҫ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���,������ȹ������ȡ�����ݽ�β
	 * @return Sheet �б�
	 * @throws ExcelIndexOutOfBoundsException ��������
	 */
	protected ArrayList<Sheet> getSheetList(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet������  
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
	 * �ж����ݳ��ȡ���ʼ�±�Ͷ�ȡ���Ȳ����Ƿ�Խ��
	 * @param count �����ܳ���,����С��1
	 * @param startIndex ��ʼ�±겻��С�����������ֵ
	 * @param length ��ȡ�ĳ���,����С��0
	 * @return �������Խ���׳��쳣,���򷵻�Ҫ��ȡ�����һ���±�ֵ(�����ȡ���ȴ����ܳ���,�򷵻�����±�ֵ)
	 * @throws ExcelIndexOutOfBoundsException ����Խ�����
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		if(count<1){ //���ݳ���С��1
			throw new ExcelIndexOutOfBoundsException();
		}
		if(length<0){//��ȡ����С����
			throw new ExcelIndexOutOfBoundsException();
        }
        if(startIndex > count -1 || startIndex < 0){//��ʼ�±���������±�ֵ��С����
        	throw new ExcelIndexOutOfBoundsException();
        }
        //Ҫ��ȡ�����һ���±�,����±�Խ�磬���ȡ�����һ��ֵ
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
}
	
	