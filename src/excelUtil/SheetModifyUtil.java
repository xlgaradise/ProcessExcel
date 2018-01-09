
package excelUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNullParameterException;

/**
*@auchor HPC
*@encoding GBK
*/

public class SheetModifyUtil {
	
	// %%%%%%%%-------字段部分 开始----------%%%%%%%%%

	/**
	 * 实例的sheet
	 */
	protected Sheet sheet;

	/**
	 * Excel工作薄
	 */
	protected Workbook workbook = null;

	// %%%%%%%%-------字段部分 结束----------%%%%%%%%%

	/**
	 * 创建写sheet的工具
	 * 
	 * @param sheet 将要写入数据的sheet,从ExcelModifyUtil获取
	 * @throws ExcelNullParameterException sheet为null
	 */
	public SheetModifyUtil(Sheet sheet) throws ExcelNullParameterException {
		if (sheet == null)
			throw new ExcelNullParameterException();
		this.sheet = sheet;
		this.workbook = sheet.getWorkbook();
	}

	public Sheet getSheet() {
		return this.sheet;
	}

	
	/**
	 * 删除某一行
	 * @param row
	 */
	public void removeRow(Row row){
		this.sheet.removeRow(row);
	}
	
	
	/**
	 * 删除某一行
	 * @param rowIndex
	 * @throws ExcelIllegalArgumentException 下标值不在有效范围
	 */
	public void removeRowAt(int rowIndex) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(rowIndex < 0 || rowIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = this.sheet.getRow(rowIndex);
		removeRow(row);
	}
	
	
	/**
	 * 删除某些行
	 * @param startIndex (base 0)
	 * @param length 
	 * @throws ExcelIllegalArgumentException 下标值不在有效范围
	 */
	public void removeRowsIn(int startIndex,int length) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(startIndex < 0 || startIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}else if((startIndex + length - 1) > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = null;
		for(int i=startIndex+length-1;i>=startIndex;i--){
			row = sheet.getRow(i);
			removeRow(row);
		}
	}
	
	/**
	 * 删除指定下标后的所有行
	 * @param startIndex
	 * @throws ExcelIllegalArgumentException 下标值不在有效范围
	 */
	public void removeRowsFrom(int startIndex) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(startIndex < 0 || startIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = null;
		for(int i=endIndex;i>=startIndex;i--){
			row = sheet.getRow(i);
			removeRow(row);
		}
	}
	
	/**
	 * 获取某一行
	 * @param rowIndex
	 * @return Row实例或null
	 */
	public Row getRowAt(int rowIndex){
		return sheet.getRow(rowIndex);
	}
	
	/**
	 * 在指定下标创建新的一行
	 * @param rowIndex
	 * @return Row实例
	 */
	public Row createNewRow(int rowIndex){
		return sheet.createRow(rowIndex);
	}
	
	/**
	 * 获取有效的Cell单元(非合并区域内部的单元)
	 * @param rowIndex 行下标
	 * @param columnIndex 列下标
	 * @return 指定单元格,或者null(无效单元格)
	 * @throws ExcelIndexOutOfBoundsException 下标参数小于零
	 * @throws IllegalArgumentException columnIndex < 0 或者 大于文件提供最大值
	 */
	public Cell getValidCell(int rowIndex,int columnIndex) throws ExcelIndexOutOfBoundsException,IllegalArgumentException{
		
		Cell cell = null;
		if (SheetReadUtil.hasMerged(sheet)) {// 如果有合并单元格
			int result = SheetReadUtil.isCellInMergedRegion(sheet,rowIndex, columnIndex);
			if (result == 1) { // 单元格是合并区域第一单元
				cell = getCell(rowIndex, columnIndex);
			} else if (result == 2) { // 单元格是合并区域内部的单元
				cell = null;
			} else { // 单元格不是合并区域内的单元
				cell = getCell(rowIndex, columnIndex);
			}
		} else {// 没有合并区域
			cell = getCell(rowIndex, columnIndex);
		}
		return cell;
		
	}
		
	/**
	 * 将数据添加至单元格中
	 * @param cell 单元格(若为null，则不执行该方法)
	 * @param value 数据值 (日期值必须符合(yyyy-MM-dd,yyyy-MM,MM-dd))
	 * @param cellType 单元格所属类型
	 * @param cellStyle 单元格样式(可为null，若cellType为日期格式则需传递新的cellStyle实例)
	 * @throws ExcelIllegalArgumentException  数据值不匹配对应格式
	 */
	public void addValueToCell(Cell cell,String value,
			TypeEnum cellType,CellStyle cellStyle) throws ExcelIllegalArgumentException{
		
		if(cell == null){
			return;
		}
		if(cellStyle == null){
			CellStyleUtil cellStyleUtil = null;
			try {
				cellStyleUtil = new CellStyleUtil(workbook);
			} catch (ExcelNullParameterException e) {
			}
			cellStyle = cellStyleUtil.getCommonCellStyle();
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
			}catch (NullPointerException |NumberFormatException e) {
				throw new ExcelIllegalArgumentException();
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
			DataFormat	dataFormat_num = sheet.getWorkbook().createDataFormat();
			cellStyle.setDataFormat(dataFormat_num.getFormat("yyyy-MM-dd"));
			Date dateNum = null;
			try {
				dateNum = new SimpleDateFormat("yyyy-MM-dd").parse(value);
			} catch (ParseException e) {
				throw new ExcelIllegalArgumentException();
			}
			cell.setCellStyle(cellStyle);
			cell.setCellValue(dateNum);
			break;
		case DATE_STR:
			DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
			short formatNum = 0;
			SimpleDateFormat simpleDateFormat = null;
			Date dateStr = null;
			//Value格式为((yyyy-MM-dd,yyyy-MM,MM-dd))
			int length = value.length();
			if(length >= 8 && length <= 10){
				simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
			 	formatNum = dataFormat.getFormat("yyyy-MM-dd");
			}else if (length == 6 || length == 7) {
				simpleDateFormat = new SimpleDateFormat("yyyy-MM");
				formatNum = dataFormat.getFormat("yyyy-MM");
			}else if(length >= 3 && length <= 5){
				simpleDateFormat = new SimpleDateFormat("MM-dd");
				formatNum = dataFormat.getFormat("MM-dd");
			}else{
				throw new ExcelIllegalArgumentException();
			}
			try {
				dateStr = simpleDateFormat.parse(value);
			} catch (ParseException e) {
				throw new ExcelIllegalArgumentException();
			}
			cellStyle.setDataFormat(formatNum);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(dateStr);
			break;
		case ERROR:
			cell.setCellStyle(cellStyle);
			//cell.setCellErrorValue(Byte.parseByte(value));
			cell.setCellValue(value);
			break;
		case FORMULA:
			cell.setCellStyle(cellStyle);
			//cell.setCellFormula(value);
			dd = 0;
			try{
				dd = Double.parseDouble(value);
				in = (int) dd;
				last = dd - in;
				if (last == 0) // double为整数
					cell.setCellValue(in);
				else
					cell.setCellValue(dd);
			}catch (NullPointerException |NumberFormatException e) {
				cell.setCellValue(value);
			}
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
	 * 设置某列自动调整列宽
	 * @param columnIndex 要调整列宽的下标
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
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
		
		CellStyleUtil cellStyleUtil = null;
		try {
			cellStyleUtil = new CellStyleUtil(workbook);
		} catch (ExcelNullParameterException e) {
		}
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
		CellStyleUtil cellStyleUtil = null;
		try {
			cellStyleUtil = new CellStyleUtil(workbook);
		} catch (ExcelNullParameterException e) {
		}
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		cellStyleUtil.setForegroundColor(cellStyle, color);
		cell.setCellStyle(cellStyle);
	}
	
	
	/**
	 * 获取指定的Cell
	 * @param rowIndex 行下标
	 * @param columnIndex 列下标
	 * @return 返回Cell,或者抛出异常
	 * @throws IllegalArgumentException  columnIndex <0 或者 大于文件提供最大值
	 */
	protected Cell getCell(int rowIndex,int columnIndex) throws IllegalArgumentException{
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		Cell cell = row.getCell(columnIndex);
		if(cell == null)
			try {
				cell = row.createCell(columnIndex);
			} catch (IllegalArgumentException e) {
				throw e;
			}
		return cell;
	}
}
