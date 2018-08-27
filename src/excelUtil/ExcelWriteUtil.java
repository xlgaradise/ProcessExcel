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
 *写Excel文件工具,仅可操作sheet
 */
public class ExcelWriteUtil {
	
	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	/**
	 * Excel文件输出目录路径
	 */
	protected String directoryPath = "";
	
	/**
	 * Excel文件名称
	 */
	protected String fileName = "";
	
	/**
	 * 操作Excel的Workbook工具
	 */
	protected SXSSFWorkbook workbook = null;
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	
	/**
	 * 创建写Excel工具
	 * @param directoryPath 文件输出所在目录路径
	 * @param fileName 文件名(不要后缀)
	 * @throws ExcelFileNotFoundException 路径无效或不是一个目录
	 * @throws ExcelIllegalArgumentException 文件名为空
	 * @throws ExcelNullParameterException directoryPath参数为null
	 * @throws SecurityException 目录文件拒绝访问
	 */
	public ExcelWriteUtil(String directoryPath,String fileName) throws ExcelFileNotFoundException,
								ExcelIllegalArgumentException,ExcelNullParameterException,SecurityException{
		if(directoryPath == null){
			throw new ExcelNullParameterException();
		}
		
		File file = new File(directoryPath);
		if (!file.exists()) // 目录是否存在
			throw new ExcelFileNotFoundException();
		if (!file.isDirectory())// 文件路径不是一个目录
			throw new ExcelFileNotFoundException();
		this.directoryPath = directoryPath;
		if (fileName == null || fileName.trim().equals(""))// 文件名为空
			throw new ExcelIllegalArgumentException();
		this.fileName = fileName;
		this.workbook = new SXSSFWorkbook(100);
	}
	
	public SXSSFWorkbook getWorkBook(){
		return workbook;
	}
	
	/**
	 * 创建新的sheet
	 * @return sheet实例
	 */
	public SXSSFSheet createSheet(){
		return this.workbook.createSheet();
	}
	
	/**
	 * 创建新的sheet
	 * @param sheetName sheet名称
	 * @return sheet实例,或者抛出异常
	 * @throws IllegalArgumentException sheet名为null,或者包含非法参数,或者名称已存在
	 */
	public SXSSFSheet createSheet(String sheetName) throws IllegalArgumentException{
		try {
			return workbook.createSheet(sheetName);
		} catch (IllegalArgumentException e) {
			throw e;
		}
	}
	
	
	/**
	 * 将workbook内容写入Excel文件
	 * @throws FileNotFoundException 如果该文件存在，但它是一个目录，而不是一个常规文件；
	 * 或者该文件不存在，但无法创建它；
	 * 抑或因为其他某些原因而无法打开它 
	 * @throws SecurityException 拒绝对文件进行写入访问
	 * @throws IOException 文件写入或者关闭出错
	 */
	public void writeToExcel() throws FileNotFoundException,SecurityException,IOException{
		
		String path = directoryPath + File.separator + fileName + ".xlsx";
		FileOutputStream outputStream = new FileOutputStream(path);
		workbook.write(outputStream);
		outputStream.close();
		workbook.dispose();
	}

	/**
	 * 关闭写入工具
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