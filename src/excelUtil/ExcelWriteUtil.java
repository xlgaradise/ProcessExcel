package excelUtil;
/**
 *@auchor HPC
 *
 */
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *дExcel�ļ�����
 */
public class ExcelWriteUtil {
	
	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	/**
	 * Excel�ļ����Ŀ¼·��
	 */
	protected String directoryPath = "";
	
	/**
	 * Excel�ļ�����
	 */
	protected String fileName = "";
	
	/**
	 *�ļ���׺��(.xls,.xlsx) ,Ĭ��Ϊ.xls
	 */
	protected String extension = ".xls";
	
	/**
	 * ����Excel��Workbook����
	 */
	protected Workbook workbook = null;
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	
	
	/**
	 * ����дExcel�ļ��Ĺ���,Ĭ���ļ���ʽΪ.xls
	 * @param directoryPath �ļ��������Ŀ¼·��
	 * @param fileName �ļ���
	 * @throws FileNotFoundException �ļ�·����Ч����һ��Ŀ¼
	 * @throws IllegalArgumentException �ļ���Ϊ��
	 * @throws NullPointerException excelPath����Ϊnull
	 * @throws SecurityException Ŀ¼�ļ��ܾ�����
	 */
	public ExcelWriteUtil(String directoryPath,String fileName) throws FileNotFoundException,
						IllegalArgumentException,NullPointerException,SecurityException{
		try {
			File file = new File(directoryPath);
			if(!file.exists()) //Ŀ¼�Ƿ����
				throw new FileNotFoundException("�ļ���Ŀ¼������");
			if(!file.isDirectory())//�ļ�·������һ��Ŀ¼
				throw new FileNotFoundException("�ļ�·������һ��Ŀ¼");
			this.directoryPath = directoryPath;
			if(fileName == null || fileName.trim().equals(""))//�ļ���Ϊ��
				throw new IllegalArgumentException("�ļ���Ϊ��");
			this.fileName = fileName;
			this.workbook = new HSSFWorkbook();
		} catch (NullPointerException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		}
	}
	
	
	/**
	 * ����дExcel����
	 * @param directoryPath �ļ��������Ŀ¼·��
	 * @param fileName �ļ���
	 * @param extension �ļ���ʽ(.xls .xlsx)
	 * @throws FileNotFoundException �ļ�·����Ч����һ��Ŀ¼
	 * @throws IllegalArgumentException �ļ���Ϊ��,���׺������
	 * @throws NullPointerException excelPath����Ϊnull
	 * @throws SecurityException Ŀ¼�ļ��ܾ�����
	 */
	public ExcelWriteUtil(String directoryPath,String fileName,String extension) throws FileNotFoundException,
						IllegalArgumentException,NullPointerException,SecurityException{
		try {
			File file = new File(directoryPath);
			if(!file.exists()) //Ŀ¼�Ƿ����
				throw new FileNotFoundException("�ļ���Ŀ¼������");
			if(!file.isDirectory())//�ļ�·������һ��Ŀ¼
				throw new FileNotFoundException("�ļ�·������һ��Ŀ¼");
			this.directoryPath = directoryPath;
			if(fileName == null || fileName.trim().equals(""))//�ļ���Ϊ��
				throw new IllegalArgumentException("�ļ���Ϊ��");
			this.fileName = fileName;
			if(extension.equals(".xls")){
				this.extension = extension;
				this.workbook = new HSSFWorkbook();
				
			}else if (extension.equals(".xlsx")) {
				this.extension = extension;
				this.workbook = new XSSFWorkbook();
			}else {
				throw new IllegalArgumentException("��׺������");
			}
		} catch (NullPointerException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		}
	}
	
	/**
	 * �����µ�sheet
	 * @return sheetʵ��
	 */
	public Sheet createSheet(){
		return this.workbook.createSheet();
	}
	
	/**
	 * �����µ�sheet
	 * @param sheetName sheet����
	 * @return sheetʵ��,�����׳��쳣
	 * @throws IllegalArgumentException sheet��Ϊnull,���߰����Ƿ�����,���������Ѵ���
	 */
	public Sheet createSheet(String sheetName) throws IllegalArgumentException{
		try {
			return workbook.createSheet(sheetName);
		} catch (IllegalArgumentException e) {
			throw e;
		}
	}
	
	
	/**
	 * ��workbook����д��Excel�ļ�
	 * @throws FileNotFoundException ������ļ����ڣ�������һ��Ŀ¼��������һ�������ļ���
	 * ���߸��ļ������ڣ����޷���������
	 * �ֻ���Ϊ����ĳЩԭ����޷����� 
	 * @throws SecurityException �ܾ����ļ�����д�����
	 * @throws IOException �ļ�д����߹رճ���
	 */
	public void writeToExcel() throws FileNotFoundException,SecurityException,IOException{
		try {
			String path = directoryPath+File.separator+fileName+extension;
			System.out.println(path);
			FileOutputStream outputStream = new FileOutputStream(path);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (FileNotFoundException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		} catch (IOException e) {
			throw e;
		}
	}

}
