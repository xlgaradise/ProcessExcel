
package excelUtil;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidOperationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.SAXException;

import exception.ExcelFileOpenException;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNotExcelFileException;
import exception.ExcelNullParameterException;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *��xml��ʽ��ȡ������excel�ļ�(����xlsx)
 */
public class ExcelReadXMLUtil {

	/**
	 * �ļ���
	 */
	protected String fileName = "";
	protected String filePath ="";
	
	protected OPCPackage opcPackage;
	protected ReadOnlySharedStringsTable sharedStringsTable;
	protected XSSFReader xssfReader;
	protected StylesTable stylesTable;
	
	/**
	 * ������ȡExcel�ļ��Ĺ���(����xlsx��ʽ)
	 * @param filePath 
	 * @throws ExcelNullParameterException ����Ϊnull
	 * @throws ExcelFileOpenException �ļ��򿪴���
	 * @throws ExcelNotExcelFileException �ļ�����xlsx��ʽ
	 */
	public ExcelReadXMLUtil(String filePath) throws ExcelNullParameterException,ExcelFileOpenException,
									ExcelNotExcelFileException {
		boolean isExcel = is07ExcelFile(filePath);
		if(isExcel){
			this.filePath = filePath;
			this.fileName = filePath.substring(filePath.lastIndexOf("\\")+1, filePath.length());
			try {
				opcPackage = OPCPackage.open(filePath, PackageAccess.READ);
				sharedStringsTable = new ReadOnlySharedStringsTable(this.opcPackage);  
		        xssfReader = new XSSFReader(this.opcPackage);   
		        stylesTable = xssfReader.getStylesTable();  
			} catch (InvalidOperationException |SAXException |IOException | OpenXML4JException e) {
				throw new ExcelFileOpenException(e);
			}
		}else{
			throw new ExcelNotExcelFileException();
		}
	}
	
	public String getFileName(){
		return fileName;
	}
	
	public String getFilePath(){
		return filePath;
	}
	
	/**
	 * ��ȡExcel�µ�sheet����
	 * @return
	 * @throws ExcelFileOpenException �ļ��򿪴���
	 */
	public int getSheetSize() throws ExcelFileOpenException{
        XSSFReader.SheetIterator iter = null;
		try {
			iter = (XSSFReader.SheetIterator) xssfReader  
			        .getSheetsData();
		} catch (InvalidFormatException | IOException e) {
			throw new ExcelFileOpenException(e);
		}  
        int index = 0;  
        while (iter.hasNext()) {  
           index++;
        }  
        return index;
	}
	
	/**
	 * ��ȡ��һ��sheet��
	 * @return ����SheetReadXMLUtilʵ��
	 * @throws ExcelFileOpenException
	 */
	public SheetReadXMLUtil readFirstSheet() throws ExcelFileOpenException{
		SheetReadXMLUtil sheetReadBigDataUtil = null;
		try {
			sheetReadBigDataUtil = readSheetByIndex(0);
		} catch (ExcelIndexOutOfBoundsException | ExcelIllegalArgumentException e) {
		}
		return sheetReadBigDataUtil;
	}
	
	/**
	 * ��ȡָ���±�ֵ��sheet��
	 * @param sheetIndex (base 0)
	 * @return ����SheetReadXMLUtilʵ��
	 * @throws ExcelFileOpenException
	 * @throws ExcelIndexOutOfBoundsException �±�С��0
	 * @throws ExcelIllegalArgumentException ָ���±�sheet������
	 */
	public SheetReadXMLUtil readSheetByIndex(int sheetIndex) throws ExcelFileOpenException,
								ExcelIndexOutOfBoundsException ,ExcelIllegalArgumentException{
		if(sheetIndex<0) throw new ExcelIndexOutOfBoundsException();
		XSSFReader.SheetIterator iter = null;
		try {
			iter = (XSSFReader.SheetIterator) xssfReader  
			        .getSheetsData();
		} catch (InvalidFormatException | IOException e) {
			throw new ExcelFileOpenException(e);
		}  
		SheetReadXMLUtil sheetReadBigDataUtil = null;
		int index = 0;
        while (iter.hasNext()) {  
        	if(sheetIndex == index){
				InputStream stream = iter.next();
				sheetReadBigDataUtil = new SheetReadXMLUtil(stylesTable, sharedStringsTable, stream);
				break;
        	}
        	index++;
        }  
        if(sheetReadBigDataUtil == null) throw new ExcelIllegalArgumentException();
		return sheetReadBigDataUtil;
	}
	
	/**
	 * ��ȡָ�����Ƶ�sheet��
	 * @param sheetName
	 * @return ����SheetReadXMLUtilʵ��
	 * @throws ExcelFileOpenException
	 * @throws ExcelIllegalArgumentException ָ������sheet������
	 */
	public SheetReadXMLUtil readSheetByName(String sheetName) throws ExcelFileOpenException,ExcelIllegalArgumentException{
		XSSFReader.SheetIterator iter = null;
		try {
			iter = (XSSFReader.SheetIterator) xssfReader  
			        .getSheetsData();
		} catch (InvalidFormatException | IOException e) {
			throw new ExcelFileOpenException(e);
		}  
		SheetReadXMLUtil sheetReadBigDataUtil = null;
        while (iter.hasNext()) {
        	InputStream stream = iter.next();
        	String name = iter.getSheetName();
        	if(name.equals(sheetName)){
				sheetReadBigDataUtil = new SheetReadXMLUtil(stylesTable, sharedStringsTable, stream);
				break;
        	}
        }  
        if(sheetReadBigDataUtil == null) throw new ExcelIllegalArgumentException();
		return sheetReadBigDataUtil;
	}
	
	public void close() throws IOException{
		opcPackage.close();
	}
	
	/**
	 * ����ļ��Ƿ�Ϊ07��Excel�ļ�
	 * @param filePath �ļ�·��
	 * @return ����ļ�Ϊxlsx��ʽ�򷵻�true,����false
	 * @throws ExcelNullParameterException �ļ�·��Ϊnull�򡰡�
	 */
	public static boolean is07ExcelFile(String filePath) throws ExcelNullParameterException{
		if(filePath == null || filePath.equals("")){
			throw new ExcelNullParameterException();
		}
		String ext = "";
		try {
			ext = filePath.substring(filePath.lastIndexOf("."));
		} catch (IndexOutOfBoundsException e) {
			
		}
		return ext.equals(".xlsx") ? true : false;
	}
}
