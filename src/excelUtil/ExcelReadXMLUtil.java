
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
 *以xml形式读取大数据excel文件(仅限xlsx)
 */
public class ExcelReadXMLUtil {

	/**
	 * 文件名
	 */
	protected String fileName = "";
	protected String filePath ="";
	
	protected OPCPackage opcPackage;
	protected ReadOnlySharedStringsTable sharedStringsTable;
	protected XSSFReader xssfReader;
	protected StylesTable stylesTable;
	
	/**
	 * 创建读取Excel文件的工具(仅限xlsx格式)
	 * @param filePath 
	 * @throws ExcelNullParameterException 参数为null
	 * @throws ExcelFileOpenException 文件打开错误
	 * @throws ExcelNotExcelFileException 文件不是xlsx格式
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
	 * 获取Excel下的sheet总数
	 * @return
	 * @throws ExcelFileOpenException 文件打开错误
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
	 * 读取第一个sheet表
	 * @return 返回SheetReadXMLUtil实例
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
	 * 读取指定下标值的sheet表
	 * @param sheetIndex (base 0)
	 * @return 返回SheetReadXMLUtil实例
	 * @throws ExcelFileOpenException
	 * @throws ExcelIndexOutOfBoundsException 下标小于0
	 * @throws ExcelIllegalArgumentException 指定下标sheet表不存在
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
	 * 读取指定名称的sheet表
	 * @param sheetName
	 * @return 返回SheetReadXMLUtil实例
	 * @throws ExcelFileOpenException
	 * @throws ExcelIllegalArgumentException 指定名称sheet表不存在
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
	 * 检查文件是否为07版Excel文件
	 * @param filePath 文件路径
	 * @return 如果文件为xlsx格式则返回true,否则false
	 * @throws ExcelNullParameterException 文件路径为null或“”
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
