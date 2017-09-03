/**
*@author HPC
*/
package excelUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import exception.ExcelIllegalArgumentException;
import exception.ExcelNullParameterException;

/**
 *�ļ��޸Ĺ���
 */
public class ExcelModifyUtil {
	
	/*%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	protected String path = "";
	
	/*%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%*/
	
	
	/**
	 * ����Excel�ļ��޸Ĺ���
	 * @param excelReadUtil
	 * @throws ExcelNullParameterException ����Ϊnull
	 */
	public ExcelModifyUtil(ExcelReadUtil excelReadUtil) throws ExcelNullParameterException{
		if(excelReadUtil == null){
			throw new ExcelNullParameterException();
		}
		this.workbook = excelReadUtil.getWorkBook();
		this.path = excelReadUtil.getFilePath();
	}
	
	public Workbook getWorkBook(){
		return workbook;
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
	 * ��ȡָ��sheet���±�
	 * @param sheetName
	 * @return index of the sheet (0 based)
	 */
	public int getSheetIndex(String sheetName){
		return this.workbook.getSheetIndex(sheetName);
	}
	
	/**
	 * ��ȡָ��sheet���±�
	 * @param sheet
	 * @return index of the sheet (0 based)
	 */
	public int getSheetIndex(Sheet sheet){
		return this.workbook.getSheetIndex(sheet);
	}
	
	/**
	 * ����ָ��sheetΪ�״̬
	 * @param sheetIndex (0-based)
	 * @throws ExcelIllegalArgumentException �±����������Ч��Χ��
	 */
	public void setActivitySheet(int sheetIndex) throws ExcelIllegalArgumentException{
		int count = this.workbook.getNumberOfSheets();
		if(sheetIndex < 0 || sheetIndex == count){
			throw new ExcelIllegalArgumentException();
		}
		this.workbook.setActiveSheet(sheetIndex);
	}
	
	/**
	 * ɾ��ָ���±��sheet
	 * @param sheetIndex (0-based)
	 * @throws ExcelIllegalArgumentException �±����������Ч��Χ��
	 */
	public void removeSheetAt(int sheetIndex) throws ExcelIllegalArgumentException{
		int count = this.workbook.getNumberOfSheets();
		if(sheetIndex < 0 || sheetIndex == count){
			throw new ExcelIllegalArgumentException();
		}
		this.workbook.removeSheetAt(sheetIndex);
	}
	
	/**
	 * ɾ��ָ�����Ƶ�sheet
	 * @param sheetName
	 * @throws ExcelIllegalArgumentException ���Ʋ���������Ч����
	 */
	public void removeSheetByName(String sheetName) throws ExcelIllegalArgumentException{
		int index = getSheetIndex(sheetName);
		removeSheetAt(index);
	}
	
	/**
	 * ��workbook�ĸ������ݱ��浽ԭExcel�ļ�
	 * @throws FileNotFoundException �ļ��ѱ���һ�������
	 * @throws SecurityException �ܾ����ļ�����д�����
	 * @throws IOException �ļ�д����߹رճ���
	 */
	public void saveExcel() throws FileNotFoundException,SecurityException,IOException{
		FileOutputStream outputStream = new FileOutputStream(this.path);
		workbook.write(outputStream);
		outputStream.flush();
		outputStream.close();
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
