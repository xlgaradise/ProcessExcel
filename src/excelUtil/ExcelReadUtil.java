package excelUtil;
/** 
 * @author HPC
 * 
 */

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;


/**
 * ��ȡExcel�ļ�,���ɻ�ȡExcel�ļ��µĸ���sheet��
 */
public class ExcelReadUtil {

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	/**
	 *�ļ���׺��(xls,xlsx) 
	 */
	private String extension = "";

	/**
	 * Excel�ļ�
	 */
	private File excelFile;
	
	/**
	 * ����Excel��Workbook����
	 */
	private Workbook workbook = null;
	
	/**
	 * ��־�������
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * ���һ�ζ�ȡ��sheet�б�
	 */
	private ArrayList<Sheet> sheetList = null; 	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * ������Excel�ļ��Ĺ���
	 * @param excelPath  Excel�ļ���ȡ·��
	 * @throws IllegalArgumentException �ļ������ڻ��ʽ����
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
	 */
	public ExcelReadUtil(String excelPath) throws IllegalArgumentException,NullPointerException,
											SecurityException{
		try {
			if(isExcelFile(excelPath)){
				this.excelFile = new File(excelPath);
				String name = this.excelFile.getName();
				this.extension = name.substring(name.lastIndexOf("."));
				FileInputStream is = new FileInputStream(excelFile); 
				//���ַ�ʽ Excel 2003/2007/2010 ���ǿ��Դ����  
		        this.workbook = WorkbookFactory.create(is) ;
			}else {
				throw new IllegalArgumentException("�ļ�����Excel�ļ�");
			}
		}catch (IllegalArgumentException e) {
			throw e;
		}catch (NullPointerException e) {
			throw e;
		}catch (SecurityException e) {
			throw e;
		}catch (Exception e) {
			logger.error("other exception in ExcelReadUtil()", e);
		}
	}
	
	/**
	 * ��ȡExcel�ļ���ʽ(.xls��.xlsx)
	 * @return
	 */
	public String getExtension() {
		return extension;
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
	 * @return ����ļ�ΪExcel��ʽ�򷵻�true,����false
	 * @throws IllegalArgumentException �ļ�������
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
	 */
	public static boolean isExcelFile(String filePath) throws IllegalArgumentException,
															NullPointerException,SecurityException{
		try {
			File file = new File(filePath);
			if(!file.exists()){
				throw new IllegalArgumentException("·������,�ļ�������");
			}else{
				String name = file.getName();
				String ext = name.substring(name.lastIndexOf("."));
				if(ext.equals(".xls") || ext.equals(".xlsx")) return true;
				else return false;
			}
		} catch (NullPointerException e) {
			throw new NullPointerException("�ļ�·��Ϊ��");
		} catch (SecurityException  e) {	
			throw new SecurityException("�ļ��ܾ�����");
		}
	}
	
	/**
	 * ��ȡ��һ��sheet
	 */
	public void readFirstSheet(){
		this.sheetList = getSheetList(0, 1);
	}
	
	/**
	 * ��ȡ�±�ΪIndex��Sheet
	 * @param index sheet���±�ֵ
	 * @throws IndexOutOfBoundsException ����Խ�����
	 */
	public void readSheetByIndex(int index) throws IndexOutOfBoundsException{
		try {
			this.sheetList = getSheetList(index, 1);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ͨ�����ƶ�ȡsheet
	 * @param name sheet������
	 * @throws IllegalArgumentException ���ƴ����޷���ȡָ��sheet
	 */
	public void readSheetByName(String name) throws IllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			sheetList.add(sheet);
		}
		else
			throw new IllegalArgumentException("�޷���ȡָ�����Ƶ�sheet");
	}
	
	/**
	 * ��ȡָ����Χ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���
	 * @throws IndexOutOfBoundsException ����Խ�����
	 */
	public void readSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		try {
			this.sheetList = getSheetList(startIndex, length);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
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
	 * ��ȡ��Ҫ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���,������ȹ������ȡ�����ݽ�β
	 * @return Sheet �б�
	 * @throws IndexOutOfBoundsException ��������
	 */
	private ArrayList<Sheet> getSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet������  
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
	 * �ж����ݳ��ȡ���ʼ�±�Ͷ�ȡ���Ȳ����Ƿ�Խ��
	 * @param count �����ܳ���,����С��1
	 * @param startIndex ��ʼ�±겻��С�����������ֵ
	 * @param length ��ȡ�ĳ���,����С��0
	 * @return �������Խ���׳��쳣,���򷵻�Ҫ��ȡ�����һ���±�ֵ(����±�Խ��,�򷵻�����±�ֵ)
	 * @throws IndexOutOfBoundsException ��������
	 */
	private int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("���ݳ���С��1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("��ȡ����С����");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("��ʼ�±���������±�ֵ��С����");
        }
        //Ҫ��ȡ�����һ���±�,����±�Խ�磬���ȡ�����һ��ֵ
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
}
	
	