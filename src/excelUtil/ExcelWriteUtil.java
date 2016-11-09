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
 *写Excel文件工具
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
	 *文件后缀名(.xls,.xlsx) ,默认为.xls
	 */
	protected String extension = ".xls";
	
	/**
	 * 操作Excel的Workbook工具
	 */
	protected Workbook workbook = null;
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	
	
	/**
	 * 创建写Excel文件的工具,默认文件格式为.xls
	 * @param directoryPath 文件输出所在目录路径
	 * @param fileName 文件名
	 * @throws FileNotFoundException 文件路径无效或不是一个目录
	 * @throws IllegalArgumentException 文件名为空
	 * @throws NullPointerException excelPath参数为null
	 * @throws SecurityException 目录文件拒绝访问
	 */
	public ExcelWriteUtil(String directoryPath,String fileName) throws FileNotFoundException,
						IllegalArgumentException,NullPointerException,SecurityException{
		try {
			File file = new File(directoryPath);
			if(!file.exists()) //目录是否存在
				throw new FileNotFoundException("文件或目录不存在");
			if(!file.isDirectory())//文件路径不是一个目录
				throw new FileNotFoundException("文件路径不是一个目录");
			this.directoryPath = directoryPath;
			if(fileName == null || fileName.trim().equals(""))//文件名为空
				throw new IllegalArgumentException("文件名为空");
			this.fileName = fileName;
			this.workbook = new HSSFWorkbook();
		} catch (NullPointerException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		}
	}
	
	
	/**
	 * 创建写Excel工具
	 * @param directoryPath 文件输出所在目录路径
	 * @param fileName 文件名
	 * @param extension 文件格式(.xls .xlsx)
	 * @throws FileNotFoundException 文件路径无效或不是一个目录
	 * @throws IllegalArgumentException 文件名为空,或后缀名错误
	 * @throws NullPointerException excelPath参数为null
	 * @throws SecurityException 目录文件拒绝访问
	 */
	public ExcelWriteUtil(String directoryPath,String fileName,String extension) throws FileNotFoundException,
						IllegalArgumentException,NullPointerException,SecurityException{
		try {
			File file = new File(directoryPath);
			if(!file.exists()) //目录是否存在
				throw new FileNotFoundException("文件或目录不存在");
			if(!file.isDirectory())//文件路径不是一个目录
				throw new FileNotFoundException("文件路径不是一个目录");
			this.directoryPath = directoryPath;
			if(fileName == null || fileName.trim().equals(""))//文件名为空
				throw new IllegalArgumentException("文件名为空");
			this.fileName = fileName;
			if(extension.equals(".xls")){
				this.extension = extension;
				this.workbook = new HSSFWorkbook();
				
			}else if (extension.equals(".xlsx")) {
				this.extension = extension;
				this.workbook = new XSSFWorkbook();
			}else {
				throw new IllegalArgumentException("后缀名错误");
			}
		} catch (NullPointerException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		}
	}
	
	/**
	 * 创建新的sheet
	 * @return sheet实例
	 */
	public Sheet createSheet(){
		return this.workbook.createSheet();
	}
	
	/**
	 * 创建新的sheet
	 * @param sheetName sheet名称
	 * @return sheet实例,或者抛出异常
	 * @throws IllegalArgumentException sheet名为null,或者包含非法参数,或者名称已存在
	 */
	public Sheet createSheet(String sheetName) throws IllegalArgumentException{
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
