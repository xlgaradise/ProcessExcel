package excelUtil;
/**
 *@auchor HPC
 *
 */


import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import excelUtil.CellTypeUtil.TypeEnum;


/**
 *写sheet表的工具
 */
public class SheetWriteUtil {
	
	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	
	/**
	 * 实例的sheet
	 */
	protected Sheet sheet;
	
	/**
	 * 该sheet的所有行数据
	 */
	//private ArrayList<Row> rowList = null;
	
	/**
	 * Excel工作薄
	 */
	protected Workbook workbook = null;
	
	//所有合并区域的极值下标
	protected int minRowIndexOfMergedRange = 0;
	protected int maxRowIndexOfMergedRange = 0;
	protected int minColumnIndexOfMergedRange = 0;
	protected int maxColumnIndexOfMergedRange = 0;
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	
	
	
	

	/**
	 * 创建写sheet的工具
	 * @param sheet 将要写入数据的sheet,从ExcelWriteUtil获取
	 * @throws IllegalArgumentException sheet为null
	 */
	public SheetWriteUtil(Sheet sheet) throws IllegalArgumentException{
		if(sheet == null)
			throw new IllegalArgumentException("sheet为null");
		this.sheet = sheet;
		this.workbook = sheet.getWorkbook();
	}
	
	
	public Sheet getSheet(){
		return this.sheet;
	}

	
	
	
	/**
	 * 添加合并单元格
	 * @param startRowIndex 起始行下标
	 * @param endRowIndex 结束行下标
	 * @param startColumnIndex 起始列下标
	 * @param endColumnIndex 结束列下标
	 * @throws IllegalArgumentException  endIndex 小于 startIndex
	 */
	public void addMergedRegion(int startRowIndex,int endRowIndex,int startColumnIndex,int endColumnIndex)
												throws IllegalArgumentException{
		if(startRowIndex > endRowIndex || startColumnIndex > endColumnIndex){
			throw new IllegalArgumentException("endIndex 小于 startIndex");
		}else{
			sheet.addMergedRegion(new CellRangeAddress(startRowIndex,endRowIndex,startColumnIndex,endColumnIndex));
			if(startRowIndex < minRowIndexOfMergedRange)
				minRowIndexOfMergedRange = startRowIndex;
			if(endRowIndex > maxRowIndexOfMergedRange)
				maxRowIndexOfMergedRange = endRowIndex;
			if(startColumnIndex < minColumnIndexOfMergedRange)
				minColumnIndexOfMergedRange = startColumnIndex;
			if(endColumnIndex > maxColumnIndexOfMergedRange)
				maxColumnIndexOfMergedRange = endColumnIndex;
		}
	}
	
	/**
	 * 获取有效行
	 * @param rowIndex 行下标
	 * @return Row实例
	 */
	public Row getValidRow(int rowIndex){
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		return row;
	}
	
	/**
	 * 获取有效的Cell单元(非合并区域内部的单元)
	 * @param rowIndex 行下标
	 * @param columnIndex 列下标
	 * @return 指定单元格,或者null(无效单元格),或者抛出异常
	 * @throws IndexOutOfBoundsException 下标参数小于零
	 * @throws IllegalArgumentException columnIndex < 0 或者 大于文件提供最大值
	 */
	public Cell getValidCell(int rowIndex,int columnIndex) throws IndexOutOfBoundsException,IllegalArgumentException{
		try {
			Cell cell = null;
			if(hasMerged()){//如果有合并单元格
				int result = isCellInMergedRegion(rowIndex, columnIndex);
				if(result == 1){ //单元格是合并区域第一单元
					cell = getCell(rowIndex, columnIndex);
				}else if(result == 2){ // 单元格是合并区域内部的单元
					cell = null;
				}else { // 单元格不是合并区域内的单元
					cell = getCell(rowIndex, columnIndex);
				}
			}else{//没有合并区域
				cell = getCell(rowIndex, columnIndex);
			}
			return cell;
		} catch (IndexOutOfBoundsException e) {
			throw e; 
		} catch (IllegalArgumentException e) {
			throw e;
		}
		
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
			DataFormat	dataFormat_num = sheet.getWorkbook().createDataFormat();
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
			DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
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
	 * 设置某列自动调整列宽
	 * @param columnIndex 要调整列宽的下标
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
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
	
	/**
	 * 判断指定的单元格是否是合并单元格  
	 * @param rowIndex 单元格行下标
	 * @param columnIndex 单元格列下标
	 * @return 1(单元格是合并区域的一个单元)、2(合并区域内部的单元)、-1(不是合并区域内的单元)
	 * @throws IndexOutOfBoundsException 参数小于零
	 */
	public int isCellInMergedRegion(int rowIndex,int columnIndex) throws IndexOutOfBoundsException{
		if (hasMerged()) {
			if(rowIndex <0 || columnIndex <0){
				throw new IndexOutOfBoundsException("参数小于零");
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(rowIndex == firstRow && columnIndex == firstColumn){//单元格为是合并区域的第一个单元
					return 1;
				}else if (rowIndex >= firstRow && rowIndex <= lastRow) {//单元格在合并区域内部
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
						return 2;
					}
				}
			}
			return -1;//单元格不在合并区域内
		}else{
			return -1;
		}
	}
	
	/**  
	* 判断sheet页中是否含有合并单元格   
	* @param sheet   
	* @return  有合并单元格返回true否则返回false
	*/  
	public boolean hasMerged() {  
	     return sheet.getNumMergedRegions() > 0 ? true : false;  
	} 

}
