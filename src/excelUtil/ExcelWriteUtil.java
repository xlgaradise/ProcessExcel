package excelUtil;
/**
 *@auchor HPC
 *
 */
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import exception.ExcelFileNotFoundException;
import exception.ExcelIllegalArgumentException;
import exception.ExcelNullParameterException;


/**
 *дExcel�ļ�����,���ɲ���sheet
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
	 * ����Excel��Workbook����
	 */
	protected SXSSFWorkbook workbook = null;
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	
	/**
	 * ����дExcel����
	 * @param directoryPath �ļ��������Ŀ¼·��
	 * @param fileName �ļ���(��Ҫ��׺)
	 * @throws ExcelFileNotFoundException ·����Ч����һ��Ŀ¼
	 * @throws ExcelIllegalArgumentException �ļ���Ϊ��
	 * @throws ExcelNullParameterException directoryPath����Ϊnull
	 * @throws SecurityException Ŀ¼�ļ��ܾ�����
	 */
	public ExcelWriteUtil(String directoryPath,String fileName) throws ExcelFileNotFoundException,
								ExcelIllegalArgumentException,ExcelNullParameterException,SecurityException{
		if(directoryPath == null){
			throw new ExcelNullParameterException();
		}
		
		File file = new File(directoryPath);
		if (!file.exists()) // Ŀ¼�Ƿ����
			throw new ExcelFileNotFoundException();
		if (!file.isDirectory())// �ļ�·������һ��Ŀ¼
			throw new ExcelFileNotFoundException();
		this.directoryPath = directoryPath;
		if (fileName == null || fileName.trim().equals(""))// �ļ���Ϊ��
			throw new ExcelIllegalArgumentException();
		this.fileName = fileName;
		this.workbook = new SXSSFWorkbook(100);
	}
	
	public SXSSFWorkbook getWorkBook(){
		return workbook;
	}
	
	/**
	 * �����µ�sheet
	 * @return sheetʵ��
	 */
	public SXSSFSheet createSheet(){
		return this.workbook.createSheet();
	}
	
	/**
	 * �����µ�sheet
	 * @param sheetName sheet����
	 * @return sheetʵ��,�����׳��쳣
	 * @throws IllegalArgumentException sheet��Ϊnull,���߰����Ƿ�����,���������Ѵ���
	 */
	public SXSSFSheet createSheet(String sheetName) throws IllegalArgumentException{
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
		
		String path = directoryPath + File.separator + fileName + ".xlsx";
		FileOutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		outputStream.close();
		workbook.dispose();
	}

	/**
	 * �ر�д�빤��
	 * @throws IOException
	 */
	public void close() throws IOException{
		try {
			this.workbook.close();
		} catch (IOException e) {
			throw e;
		}
	}
}