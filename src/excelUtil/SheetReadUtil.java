package excelUtil;
/** 
 * @author HPC
 * 
 */
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class SheetReadUtil {

	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	
	/**
	 * 实例的sheet
	 */
	private Sheet sheet;
	
	/**
	 * 该sheet的所有行数据
	 */
	private ArrayList<Row> allRowList = null;
	
	/**
	 * 自选标题(Integer为标题的列下标,String为标题内容)
	 */
	private HashMap<Integer, String> titles;
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	

	/**
	 * 创建读sheet的工具
	 * @param sheet 将要操作的sheet实例,从ExcelReadUtil获取
	 */
	public SheetReadUtil(Sheet sheet){
		this.sheet = sheet;
		allRowList = new ArrayList<>();
		titles = new HashMap<>();
	}
	

	/**
	 * 获取实例的sheet
	 * @return 返回该SheetReadUtil操作的sheet
	 */
	public Sheet getSheet(){
		return sheet;
	}
	
	/**
	 * 读取Sheet中所有的行
	 */
	public void readAllRows(){
		int rowsCount = sheet.getLastRowNum() + 1;
		readRows(0, rowsCount);
	}

	/**
	 * 读取sheet中指定的行
	 * @param startIndex 开始的行下标
	 * @param length 读取长度
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	public void readRows(int startIndex,int length) throws IndexOutOfBoundsException{
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
			for(int i = startIndex;i<=endIndex;i++){
				if(endIndex == 0){//只读sheet的第一行
					Row r = sheet.getRow(0);
					if(r != null)
						this.allRowList.add(r);
				}else{
					this.allRowList.add(sheet.getRow(i));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 返回allRowList列表
	 * @return
	 */
	public ArrayList<Row> getAllRowList() {
		return allRowList;
	}
	
	/**
	 * 获取allRowList列表里指定的Row
	 * @param rowIndex Row下标值
	 * @return 返回指定Row,或者抛出异常
	 * @throws IndexOutOfBoundsException 下标值越界异常 
	 */
	public Row getRowAt(int rowIndex) throws IndexOutOfBoundsException {
		try {
			return allRowList.get(rowIndex);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 通过指定标题属性值从数据集里获取指定的RowList列表
	 * @param rows Row列表数据集
	 * @param title 标题名称
	 * @param value 标题属性值
	 * @return 返回指定RowList列表
	 * @throws IllegalArgumentException 标题值不存在
	 */
	public ArrayList<Row> getRowListByArg(ArrayList<Row> rows,String title,String value) throws IllegalArgumentException{
		int columnIndex = getTitleColIndexByValue(title);
		ArrayList<Row> rowList = null;
		if(columnIndex != -1){//title值存在
			rowList = new ArrayList<>();
			for(Row row : rows){
				String v = getCellValue(row.getCell(columnIndex));
				if(v.equals(value)){
					rowList.add(row);
				}
			}
			return rowList;
		}else{
			throw new IllegalArgumentException("title值不存在");
		}
	}
	
	/**
	 * 通过指定标题属性值从数据集获取指定的RowList列表
	 * @param startRowIndex 所需数据集在allRowList里的起始下标
	 * @param length 所需数据集的长度
	 * @param title 标题名称
	 * @param value 标题属性值
	 * @return 返回指定RowList列表
	 * @throws IndexOutOfBoundsException 下标值越界
	 * @throws IllegalArgumentException 标题值不存在
	 */
	public ArrayList<Row> getRowListByArg(int startRowIndex,int length,String title,String value) throws
													IndexOutOfBoundsException,IllegalArgumentException{
		int count = allRowList.size();
		try {
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
			ArrayList<Row> rows = new ArrayList<>();
			for(int i=startRowIndex;i<=endIndex;i++){
				rows.add(allRowList.get(i));
			}
			return getRowListByArg(rows, title, value);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}catch (IllegalArgumentException e) {
			throw e;
		}
	}
	
	/**
	 * 获取指定Row的所有CellList
	 * @param row 指定Row
	 * @return 返回cellList
	 */
	public ArrayList<Cell> getOneRowAllCells(Row row){
		int count = row.getLastCellNum();
		if(count == -1){
			return new ArrayList<Cell>();
		}
		return getOneRowCellList(row, 0, count);
	}
	
	/**
	 * 获取指定Row中指定的CellList
	 * @param row 指定Row
	 * @param startColumnIndex 起始的列下标值
	 * @param length 所需长度值
	 * @return 返回cellList或者抛出错误
	 * @throws IndexOutOfBoundsException 下标值越界
	 */
	public ArrayList<Cell> getOneRowCellList(Row row,int startColumnIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = row.getLastCellNum();
		if(count == -1){
			return cellList;
		}
		try {
			int endIndex = isIndexOutOfBounds(count, startColumnIndex, length);
			Cell c = null;
			for(int i=startColumnIndex;i<=endIndex;i++){
				c = row.getCell(i);
				cellList.add(c);
			}
			return cellList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 获取一列的所有Cells
	 * @param columnIndex 指定的列下标值
	 * @return 返回cellList或者抛出错误
	 * @throws IndexOutOfBoundsException 下标值越界
	 */
	public ArrayList<Cell> getOneColumnAllCells(int columnIndex) throws IndexOutOfBoundsException{
		try {
			int length = allRowList.size();
			return getOneColumnCellList(columnIndex, 0, length);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 获取一列中指定的Cells
	 * @param cloumnIndex 指定的列下标值
	 * @param startRowIndex 开始的行下标值
	 * @param length 所需长度
	 * @return 返回cellList列表或者抛出错误
	 * @throws IndexOutOfBoundsException 下标值越界
	 */
	public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = allRowList.size();
		try {
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
			Cell cell = null;
			Row row = null;
			for(int i=startRowIndex;i<=endIndex;i++){
				row = allRowList.get(i);
				if(row == null){
					cellList.add(null);
				}else{
					cell = row.getCell(cloumnIndex);
					cellList.add(cell);
				}
			}
			return cellList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 设定sheet的标题列表
	 * @param row 指定的行
	 * @param startColumnIndex 开始列下标值
	 * @param length 所需长度
	 * @throws IndexOutOfBoundsException 下标值越界
	 * @throws IllegalArgumentException 指定行没有数据
	 */
	public void setTitles(Row row,int startColumnIndex,int length) throws IndexOutOfBoundsException,
																IllegalArgumentException{
		int cellCount = row.getLastCellNum();
		if(cellCount == -1){
			throw new IllegalArgumentException("该行没有数据");
		}
		try {
			ArrayList<Cell> cells = getOneRowCellList(row, startColumnIndex, length);
			Cell cell = null;
			for(int i=0;i<cells.size();i++){
				cell = cells.get(i);
				if(cell == null){
					titles.put(startColumnIndex+i,"");
				}else{
					titles.put(startColumnIndex+i, getCellValue(cell));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 获取标题列表
	 * @return 返回标题
	 */
	public HashMap<Integer, String> getTitles(){
		return titles;
	}
	
	/**
	 * 通过标题值获取所在列下标
	 * @param titleName 指定标题值
	 * @return 标题列下标,不存在则返回-1
	 */
	public int getTitleColIndexByValue(String titleName){
		for(Map.Entry<Integer, String> entry : titles.entrySet()){
			if(entry.getValue().equals(titleName)) 
				return entry.getKey();
		}
		return -1;
	}
	
	/**
	 * 获取指定Cell的数据值
	 * @param cell 指定Cell
	 * @return 将所有数据以String类型返回
	 */
	@SuppressWarnings("deprecation")
	public String getCellValue(Cell cell){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			CellType cellType = cell.getCellTypeEnum();
			
			switch (cellType) {
			case STRING:
				string =  cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				short format = cell.getCellStyle().getDataFormat();
				/**
				 * 判断日期类型
				 * (2001.01.01) cellType:STRING format：0
				 * (yyyy-MM) cellType:NUM format:17 是日期格式
				 * (yyyy年MM月dd日) cellType:NUM format:31
				 * (yyyy年MM月) cellType:NUM format:57
				 * (MM-dd或MM月dd日) cellType:NUM format:58
				 * (yyyy-MM-dd) cellType:NUM format: 176 
				 * 177? 数值与日期均存在
				 * 179?
				 * 178?
				 */
				if(DateUtil.isCellDateFormatted(cell)){ //处理日期格式
					SimpleDateFormat s = new SimpleDateFormat("yyyy-MM-dd");
					double d = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(d);
					string = s.format(date);
				}else if(format == 31 || format == 57 || format == 58 || format == 176){//自定义日期格式
					SimpleDateFormat s = null;
					switch (format) {
					case 31:
						s = new SimpleDateFormat("yyyy年MM月dd日");
						break;
					case 57:
						s = new SimpleDateFormat("yyyy年MM月");
						break;
					case 58:
						s = new SimpleDateFormat("MM月dd日");
						break;
					case 176:
						s = new SimpleDateFormat("yyyy年MM月dd日");
						break;
					default:
						break;
					}
					double d = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(d);
					string = s.format(date);
				}else{
					double d = cell.getNumericCellValue();
					int in = (int)d;
					double last = d - in;
					if(last == 0) //double为整数
						string = String.valueOf(in);
					else 
						string = String.valueOf(d);
				}
				break;
			case ERROR:
				string = String.valueOf(cell.getErrorCellValue());
				break;
			case FORMULA:
				string = cell.getCellFormula();
				break;
			case BOOLEAN:
				string = String.valueOf(cell.getBooleanCellValue());
				break;
			case BLANK:
				string = "";
				break;
			default:
				string = "";
			}
			return string.trim();
		} catch (Exception e) {//未知错误
			return "";
		}
	}
	
	/**
	 * 获取合并单元格的值 
	 * @param cell 指定单元格
	 * @return 单元格的值
	 * @throws IllegalArgumentException 参数cell为null,或者cell不是合并单元格
	 */
	public String getCellValueOfMergedRegion(Cell cell) throws IllegalArgumentException{  
		try {
			int result = isCellInMergedRegion(cell);
			if (result == 2) { //单元格在合并区域内部
				int sheetMergeCount = sheet.getNumMergedRegions();
				int rowIndex = cell.getRowIndex();
				int columnIndex = cell.getColumnIndex();
				for (int i = 0; i < sheetMergeCount; i++) {
					CellRangeAddress ca = sheet.getMergedRegion(i);
					int firstColumn = ca.getFirstColumn();
					int lastColumn = ca.getLastColumn();
					int firstRow = ca.getFirstRow();
					int lastRow = ca.getLastRow();
					
					if (rowIndex >= firstRow && rowIndex <= lastRow) {
						if (columnIndex >= firstColumn && columnIndex <= lastColumn) { //确定区域位置
							Row fRow = sheet.getRow(firstRow);
							Cell fCell = fRow.getCell(firstColumn);
							return getCellValue(fCell); //返回该区域的第一个单元格的值
						}
					}
				}
				return "";
			}else if (result == 1) { //单元格为合并区域的第一个单元
				return getCellValue(cell); //直接返回该单元格的值
			}else{
				throw new IllegalArgumentException("指定单元格不是合并单元格");
			}   
		} catch (IllegalArgumentException e) {
			throw e;
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
	
	/**
	 * 判断指定的单元格是否是合并单元格  
	 * @param cell 指定单元格
	 * @return 1(单元格是合并区域的一个单元)、2(合并区域内部的单元)、-1(不是合并区域内的单元)
	 * @throws IllegalArgumentException 参数为null
	 */
	public int isCellInMergedRegion(Cell cell) throws IllegalArgumentException{
		if (hasMerged()) {
			if(cell == null){
				throw new IllegalArgumentException("参数为null");
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			int rowIndex = cell.getRowIndex();
			int columnIndex = cell.getColumnIndex();
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
			return -1;
		}else{
			return -1;
		}
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
	 * 判断数据长度、起始下标和读取长度参数是否越界
	 * @param count 数据总长度,不能小于1
	 * @param startIndex 起始下标不能小于零或大于最大值
	 * @param length 读取的长度,不能小于0
	 * @return 如果参数越界抛出异常,否则返回要读取的最后一个下标值(如果下标越界,则返回最大下标值)
	 * @throws IndexOutOfBoundsException 参数错误
	 */
	private int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("数据长度小于1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("读取长度小于零");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("开始下标大于最大的下标值或小于零");
        }
        //要读取的最后一个下标,如果下标越界，则读取至最后一个值
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
	
}
