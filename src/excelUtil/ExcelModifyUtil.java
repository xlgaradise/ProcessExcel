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
 *文件修改工具
 */
public class ExcelModifyUtil {
	
	/*%%%%%%%%-------字段部分 开始----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	protected String path = "";
	
	/*%%%%%%%%-------字段部分 结束----------%%%%%%%%%*/
	
	
	/**
	 * 创建Excel文件修改工具
	 * @param excelReadUtil
	 * @throws ExcelNullParameterException 参数为null
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
	 * 获取指定sheet的下标
	 * @param sheetName
	 * @return index of the sheet (0 based)
	 */
	public int getSheetIndex(String sheetName){
		return this.workbook.getSheetIndex(sheetName);
	}
	
	/**
	 * 获取指定sheet的下标
	 * @param sheet
	 * @return index of the sheet (0 based)
	 */
	public int getSheetIndex(Sheet sheet){
		return this.workbook.getSheetIndex(sheet);
	}
	
	/**
	 * 设置指定sheet为活动状态
	 * @param sheetIndex (0-based)
	 * @throws ExcelIllegalArgumentException 下标参数不在有效范围内
	 */
	public void setActivitySheet(int sheetIndex) throws ExcelIllegalArgumentException{
		int count = this.workbook.getNumberOfSheets();
		if(sheetIndex < 0 || sheetIndex == count){
			throw new ExcelIllegalArgumentException();
		}
		this.workbook.setActiveSheet(sheetIndex);
	}
	
	/**
	 * 删除指定下标的sheet
	 * @param sheetIndex (0-based)
	 * @throws ExcelIllegalArgumentException 下标参数不在有效范围内
	 */
	public void removeSheetAt(int sheetIndex) throws ExcelIllegalArgumentException{
		int count = this.workbook.getNumberOfSheets();
		if(sheetIndex < 0 || sheetIndex == count){
			throw new ExcelIllegalArgumentException();
		}
		this.workbook.removeSheetAt(sheetIndex);
	}
	
	/**
	 * 删除指定名称的sheet
	 * @param sheetName
	 * @throws ExcelIllegalArgumentException 名称参数不是有效名称
	 */
	public void removeSheetByName(String sheetName) throws ExcelIllegalArgumentException{
		int index = getSheetIndex(sheetName);
		removeSheetAt(index);
	}
	
	/**
	 * 将workbook的跟新内容保存到原Excel文件
	 * @throws FileNotFoundException 文件已被另一个程序打开
	 * @throws SecurityException 拒绝对文件进行写入访问
	 * @throws IOException 文件写入或者关闭出错
	 */
	public void saveExcel() throws FileNotFoundException,SecurityException,IOException{
		FileOutputStream outputStream = new FileOutputStream(this.path);
		workbook.write(outputStream);
		outputStream.flush();
		outputStream.close();
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
