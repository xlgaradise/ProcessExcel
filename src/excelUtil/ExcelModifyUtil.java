/**
*@author HPC
*/
package excelUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import excelUtil.CellTypeUtil.TypeEnum;

public class ExcelModifyUtil {
	
	/*%%%%%%%%-------字段部分 开始----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	protected String path = "";
	
	/*%%%%%%%%-------字段部分 结束----------%%%%%%%%%*/
	
	
	/**
	 * 创建Excel文件修改工具
	 * @param excelReadUtil
	 * @throws IllegalArgumentException 参数为null
	 */
	public ExcelModifyUtil(ExcelReadUtil excelReadUtil) throws IllegalArgumentException{
		if(excelReadUtil == null){
			throw new IllegalArgumentException();
		}
		this.workbook = excelReadUtil.getWorkBook();
		this.path = excelReadUtil.getFilePath();
	}
	
	
	/**
	 * 将数据添加至单元格中
	 * @param cell 单元格(若为null，则不执行该方法)
	 * @param value 数据值 (日期值必须符合(yyyy-MM-dd,yyyy-MM,MM-dd))
	 * @param cellType 单元格所属类型
	 * @param cellStyle 单元格样式(可为null，若cellType为日期格式则需传递新的cellStyle实例)
	 * @throws (-----详细信息保存在message里-------)
	 * @throws IllegalArgumentException  数据值格式错误
	 */
	public void addValueToCell(Cell cell,String value,
			TypeEnum cellType,CellStyle cellStyle) throws IllegalArgumentException{
		
		if(cell == null){
			return;
		}
		if(cellStyle == null){
			CellStyleUtil cellStyleUtil = new CellStyleUtil(workbook);
			cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		}
		switch (cellType) {
		case STRING:
			cell.setCellStyle(cellStyle);
			cell.setCellValue(value);
			break;
		case NUMERIC:
			double dd = 0;
			try{
				dd = Double.parseDouble(value);
			}catch (NullPointerException e) {
				throw new IllegalArgumentException("numeric value can't parse to double");
			}catch (NumberFormatException e) {
				throw new IllegalArgumentException("numeric value can't parse to double");
			}
			int in = (int) dd;
			double last = dd - in;
			cell.setCellStyle(cellStyle);
			if (last == 0) // double为整数
				cell.setCellValue(in);
			else
				cell.setCellValue(dd);
			break;
		case DATE_NUM:
			DataFormat	dataFormat_num = this.workbook.createDataFormat();
			cellStyle.setDataFormat(dataFormat_num.getFormat("yyyy-MM-dd"));
			Date dateNum = null;
			try {
				dateNum = new SimpleDateFormat("yyyy-MM-dd").parse(value);
			} catch (ParseException e) {
				throw new IllegalArgumentException("Date value is not dateFormat");
			}
			cell.setCellStyle(cellStyle);
			cell.setCellValue(dateNum);
			break;
		case DATE_STR:
			DataFormat dataFormat = this.workbook.createDataFormat();
			short formatNum = 0;
			SimpleDateFormat simpleDateFormat = null;
			Date dateStr = null;
			//Value格式为((yyyy-MM-dd,yyyy-MM,MM-dd))
			int length = value.length();
			if(length == 10){
				simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
			 	formatNum = dataFormat.getFormat("yyyy-MM-dd");
			}else if (length == 7) {
				simpleDateFormat = new SimpleDateFormat("yyyy-MM");
				formatNum = dataFormat.getFormat("yyyy-MM");
			}else {
				simpleDateFormat = new SimpleDateFormat("MM-dd");
				formatNum = dataFormat.getFormat("MM-dd");
			}
			try {
				dateStr = simpleDateFormat.parse(value);
			} catch (ParseException e) {
				throw new IllegalArgumentException("Date value is not dateFormat");
			}
			cellStyle.setDataFormat(formatNum);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(dateStr);
			break;
		case ERROR:
			cell.setCellStyle(cellStyle);
			cell.setCellErrorValue(Byte.parseByte(value));
			break;
		case FORMULA:
			cell.setCellStyle(cellStyle);
			cell.setCellFormula(value);
			break;
		case BOOLEAN:
			cell.setCellStyle(cellStyle);
			cell.setCellValue(Boolean.parseBoolean(value));
			break;
		case BLANK:
		default:
			cell.setCellStyle(cellStyle);
			cell.setCellValue("");
		}
	}
	

	/**
	 * 将workbook内容写入Excel文件
	 * @throws FileNotFoundException 文件已被另一个程序打开
	 * @throws SecurityException 拒绝对文件进行写入访问
	 * @throws IOException 文件写入或者关闭出错
	 */
	public void writeToExcel() throws FileNotFoundException,SecurityException,IOException{
		try {
			FileOutputStream outputStream = new FileOutputStream(this.path);
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
	
	/**
	 * 设置指定行的前景色
	 * @param row Row实例(如果Row为null或者没有Cell则不执行方法)
	 * @param color 颜色
	 */
	public void setForegroundColor(Row row,IndexedColors color){
		if(row == null){
			return;
		}
		int count = row.getLastCellNum();
		if(count == -1){
			return;
		}
		
		CellStyleUtil cellStyleUtil = new CellStyleUtil(workbook);
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		for(int i=0;i<count;i++){
			Cell cell = row.getCell(i);
			if(cell == null){
				cell = row.createCell(i);
			}
			cellStyleUtil.setForegroundColor(cellStyle, color);
			cell.setCellStyle(cellStyle);
		}
	}
	
	/**
	 * 设置指定单元格的前景色
	 * @param cell 指定单元格(如果Cell为null则不执行方法)
	 * @param color 指定颜色
	 */
	public void setForegroundColor(Cell cell,IndexedColors color){
		if(cell == null)
			return;
		CellStyleUtil cellStyleUtil = new CellStyleUtil(workbook);
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		cellStyleUtil.setForegroundColor(cellStyle, color);
		cell.setCellStyle(cellStyle);
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
