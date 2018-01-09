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

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNoTitleException;
import exception.ExcelNullParameterException;

/**
 *读sheet表的工具,可获取行、列、单元格等
 */
public class SheetReadUtil {

	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	
	/**
	 * 实例的sheet
	 */
	protected Sheet sheet;
	
	/**
	 * 该sheet的指定行数据
	 */
	protected ArrayList<Row> rowList = null;
	
	/**
	 * 自选标题(Integer为标题的列下标,String为标题内容)
	 */
	protected HashMap<Integer, String> titles;
	
	
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	

	/**
	 * 创建读sheet的工具（从ExcelReadUtil获取实例）
	 * @param sheet 将要操作的sheet实例,从ExcelReadUtil获取
	 * @throws ExcelNullParameterException 参数为null
	 */
	public SheetReadUtil(Sheet sheet) throws ExcelNullParameterException{
		if(sheet == null){
			throw new ExcelNullParameterException();
		}
		this.sheet = sheet;
		rowList = new ArrayList<>();
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
	 * 获取sheet中有效的行数
	 * @return
	 */
	public int getPhysicalNumberOfRows(){
		return sheet.getPhysicalNumberOfRows();
	}
	
	/**
	 * 读取Sheet中所有的行
	 */
	public void readAllRows(){
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			readRows(0, rowsCount);
		} catch (ExcelIndexOutOfBoundsException e) {
		}
	}

	/**
	 * 读取sheet中指定的行
	 * @param startIndex 开始的行下标
	 * @param length 读取长度
	 * @throws ExcelIndexOutOfBoundsException 起始下标不能小于零或大于最大值，length读取的长度不能小于0
	 */
	public void readRows(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
		for (int i = startIndex; i <= endIndex; i++) {
			if (endIndex == 0) {// 只读sheet的第一行
				Row r = sheet.getRow(0);
				if (r != null)
					this.rowList.add(r);
			} else {
				this.rowList.add(sheet.getRow(i));
			}
		}
	}
	
	/**
	 * 返回rowList列表
	 * @return
	 */
	public ArrayList<Row> getRowList() {
		return rowList;
	}
	
	
	
	/**
	 * 通过指定标题属性值从数据集里获取指定的RowList列表
	 * @param rows Row列表数据集
	 * @param title 标题名称
	 * @param value 标题属性值
	 * @return 返回指定RowList列表
	 * @throws ExcelIllegalArgumentException 指定标题不存在
	 * @throws ExcelNoTitleException 未设置标题
	 */
	/*public ArrayList<Row> getRowListByArg(ArrayList<Row> rows,String title,String value) 
			throws ExcelIllegalArgumentException,ExcelNoTitleException{
		
		int columnIndex = getTitleColIndexByValue(title);
		ArrayList<Row> rowList = null;
		if(columnIndex != -1){//title值存在
			rowList = new ArrayList<>();
			for(Row row : rows){
				if(row == null){
					continue;
				}
				String v = getCellValue(row.getCell(columnIndex));
				if(v.equals(value)){
					rowList.add(row);
				}
			}
			return rowList;
		}else{
			throw new ExcelIllegalArgumentException();
		}
	}*/
	
	/**
	 * 通过指定标题属性值从数据集获取指定的RowList列表
	 * @param startRowIndex 所需数据集在allRowList里的起始下标
	 * @param length 所需数据集的长度
	 * @param title 标题名称
	 * @param value 标题属性值
	 * @return 返回指定RowList列表
	 * @throws ExcelIndexOutOfBoundsException 起始下标不能小于零或大于最大值，length读取的长度不能小于0
	 * @throws ExcelIllegalArgumentException 指定标题不存在
	 * @throws ExcelNoTitleException 未设置标题
	 */
	/*public ArrayList<Row> getRowListByArg(int startRowIndex,int length,String title,String value) throws
				ExcelIndexOutOfBoundsException,ExcelIllegalArgumentException,ExcelNoTitleException{
		int count = rowList.size();
		ArrayList<Row> returnList = new ArrayList<>();
		
		int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
		ArrayList<Row> rows = new ArrayList<>();
		for (int i = startRowIndex; i <= endIndex; i++) {
			rows.add(rowList.get(i));
		}
		returnList = getRowListByArg(rows, title, value);
		return returnList;
	}*/
	
	
	
	/**
	 * 获取指定Row的所有CellList
	 * @param row 指定Row
	 * @return 返回cellList
	 * @throws ExcelNullParameterException 参数为null
	 */
	/*public ArrayList<Cell> getOneRowAllCells(Row row) throws ExcelNullParameterException{
		if(row == null){
			throw new ExcelNullParameterException();
		}
		int count = row.getLastCellNum();
		if(count == -1){
			return new ArrayList<Cell>();
		}
		try {
			return getOneRowCellList(row, 0, count);
		} catch (ExcelIndexOutOfBoundsException e) {
			return new ArrayList<Cell>();
		}
	}*/
	
	/**
	 * 获取指定Row中指定的CellList
	 * @param row 指定Row
	 * @param startColumnIndex 起始的列下标值
	 * @param length 所需长度值
	 * @return 返回cellList或者抛出错误
	 * @throws ExcelIndexOutOfBoundsException 起始下标不能小于零或大于最大值，length读取的长度不能小于0
	 * @throws ExcelNullParameterException 参数row为null
	 */
	public ArrayList<Cell> getOneRowCellList(Row row,int startColumnIndex,int length) throws 
					ExcelIndexOutOfBoundsException,ExcelNullParameterException{
		if(row == null){
			throw new ExcelNullParameterException();
		}
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = row.getLastCellNum();
		if(count == -1){
			return cellList;
		}
		
		int endIndex = isIndexOutOfBounds(count, startColumnIndex, length);
		Cell c = null;
		for (int i = startColumnIndex; i <= endIndex; i++) {
			c = row.getCell(i);
			cellList.add(c);
		}
		return cellList;
		
	}
	
	/**
	 * 获取一列的所有Cells
	 * @param columnIndex 指定的列下标值
	 * @return 返回cellList或者抛出错误
	 * @throws ExcelIndexOutOfBoundsException 下标值越界
	 */
	/*public ArrayList<Cell> getOneColumnAllCells(int columnIndex) throws ExcelIndexOutOfBoundsException{
		int length = rowList.size();
		return getOneColumnCellList(columnIndex, 0, length);
	}*/
	
	/**
	 * 获取一列中指定的Cells
	 * @param cloumnIndex 指定的列下标值
	 * @param startRowIndex 开始的行下标值
	 * @param length 所需长度
	 * @return 返回cellList列表或者抛出错误
	 * @throws ExcelIndexOutOfBoundsException 起始下标不能小于零或大于最大值，length读取的长度不能小于0
	 */
	/*public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) 
									throws ExcelIndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = rowList.size();
		
		int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
		Cell cell = null;
		Row row = null;
		for (int i = startRowIndex; i <= endIndex; i++) {
			row = rowList.get(i);
			if (row == null) {
				cellList.add(null);
			} else {
				cell = row.getCell(cloumnIndex);
				cellList.add(cell);
			}
		}
		return cellList;
	}*/
	
	/**
	 * 设定sheet的标题列表
	 * @param rowIndex sheet中行下标
	 * @param startColumnIndex 开始列下标值
	 * @param length 所需长度
	 * @throws ExcelIndexOutOfBoundsException 起始下标不能小于零或大于最大值，length读取的长度不能小于0
	 * @throws ExcelIllegalArgumentException 指定下标行为null,或没有数据
	 */
	public void setTitle(int rowIndex,int startColumnIndex,int length) throws ExcelIndexOutOfBoundsException,
													ExcelIllegalArgumentException{
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			throw new ExcelIllegalArgumentException();
		int cellCount = row.getLastCellNum();
		if(cellCount == -1){
			throw new ExcelIllegalArgumentException();
		}
		titles.clear();
		
		ArrayList<Cell> cells = null;
		try {
			cells = getOneRowCellList(row, startColumnIndex, length);
		} catch (ExcelNullParameterException e) {
		}
		Cell cell = null;
		for (int i = 0; i < cells.size(); i++) {
			cell = cells.get(i);
			if (cell == null) {
				titles.put(startColumnIndex + i, "");
			} else {
				titles.put(startColumnIndex + i, getCellValue(cell));
			}
		}
	}
	
	/**
	 * 获取标题列表
	 * @return 返回标题
	 */
	public HashMap<Integer, String> getTitle(){
		return titles;
	}
	
	/**
	 * 通过标题值获取所在列下标
	 * @param titleName 指定标题值
	 * @return 标题列下标,不存在则返回-1
	 * @throws ExcelNoTitleException 未设置标题
	 */
	public int getTitleColIndexByValue(String titleName) throws ExcelNoTitleException{
		if(titles.isEmpty()){
			throw new ExcelNoTitleException();
		}
		for(Map.Entry<Integer, String> entry : titles.entrySet()){
			if(entry.getValue().equals(titleName)) 
				return entry.getKey();
		}
		return -1;
	}
	
	/**
	 * 获取指定Cell的数据值
	 * @param cell 指定Cell
	 * @return 将所有数据以String类型返回;
	 * <br>日期类型格式(yyyy-MM-dd,yyyy-MM,MM-dd),单元格式为DATE_NUM时只返回(yyyy-MM-dd);
	 * <br>公式类型格式返回公式计算方法而不是值;
	 * <br>cell为null,或没有值,或取值出错则返回""。
	 */
	public static String getCellValue(Cell cell){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			TypeEnum cellType = CellTypeUtil.getCellType(cell);
			
			switch (cellType) {
			case STRING:
				string =  cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				double dd = cell.getNumericCellValue();
				int in = (int) dd;
				double last = dd - in;
				if (last == 0) // double为整数
					string = String.valueOf(in);
				else
					string = String.valueOf(dd);
				break;
			case DATE_NUM:
				double d = cell.getNumericCellValue();
				string = CellTypeUtil.getFormatDate(d);
				break;
			case DATE_STR:
				string =  cell.getStringCellValue().trim();
				string = CellTypeUtil.getFormatDate(string);
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
			//去掉首尾全角空白符
			while (string.startsWith("　")) {
				string = string.substring(1, string.length()).trim();
			}
			while (string.endsWith("　")) {
				string = string.substring(0, string.length() - 1).trim();
			}
			/*if (string != null) {//去掉换行、回车、制表符  
		        Pattern p = Pattern.compile("\\s*|\t|\r|\n");  
		        Matcher m = p.matcher(string);  
		        string = m.replaceAll("");  
		    }  */
			return string.trim();
		} catch (Exception e) {//未知错误
			return "";
		}
	}
	
	/**
	 * 获取指定Cell的数据值
	 * @param cell 指定Cell
	 * @param evaluator 用于读取公式值的FormulaEvaluator实例
	 * @return 将所有数据以String类型返回;
	 * <br>日期类型格式(yyyy-MM-dd,yyyy-MM,MM-dd),单元格式为DATE_NUM时只返回(yyyy-MM-dd);
	 * <br>公式类型格式返回公式计算值(若evaluator为null则返回公式);
	 * <br>cell为null,或没有值,或取值出错则返回""。
	 */
	public static String getCellValue(Cell cell,FormulaEvaluator evaluator){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			TypeEnum cellType = CellTypeUtil.getCellType(cell);
			
			switch (cellType) {
			case FORMULA:
				if(evaluator == null){
					string = cell.getCellFormula();
				}else{
					string = getFormulaValue(evaluator.evaluate(cell));
				}
				break;
			default:
				string = getCellValue(cell);
			}
			//去掉首尾全角空白符
			while (string.startsWith("　")) {
				string = string.substring(1, string.length()).trim();
			}
			while (string.endsWith("　")) {
				string = string.substring(0, string.length() - 1).trim();
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
	 * @rhrows ExcelNullParameterException 参数cell为null
	 * @throws ExcelIllegalArgumentException cell不是合并单元格
	 */
	public static String getCellValueOfMergedRegion(Sheet sheet,Cell cell) throws ExcelNullParameterException,ExcelIllegalArgumentException{  
		if(cell == null){
			throw new ExcelNullParameterException();
		}
		
		int result = isCellInMergedRegion(sheet,cell);
		if (result == 2) { // 单元格在合并区域内部
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
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) { // 确定区域位置
						Row fRow = sheet.getRow(firstRow);
						Cell fCell = fRow.getCell(firstColumn);
						return getCellValue(fCell); // 返回该区域的第一个单元格的值
					}
				}
			}
			return "";
		} else if (result == 1) { // 单元格为合并区域的第一个单元
			return getCellValue(cell); // 直接返回该单元格的值
		} else {
			throw new ExcelIllegalArgumentException();
		}
		
	}    
	
	/**
	 * 强制将Numeric类型的值转换为日期
	 * @param numeric numeric格式的单元格数值
	 * @return yyyy-MM-dd格式的日期
	 * @throws NullPointerException 日期强转出错
	 */
	public static String changeNumericToDate(double numeric) throws NullPointerException{
		String string = "";
		Date date = DateUtil.getJavaDate(numeric);
		if(date == null){
			date = HSSFDateUtil.getJavaDate(numeric);
		}
		string = new SimpleDateFormat("yyyy-MM-dd").format(date);
		return string;
	}
	
	/**  
	* 判断sheet页中是否含有合并单元格   
	* @param sheet   
	* @return  有合并单元格返回true否则返回false
	*/  
	public static boolean hasMerged(Sheet sheet) {  
	     return sheet.getNumMergedRegions() > 0 ? true : false;  
	} 
	
	/**
	 * 判断指定的单元格是否是合并单元格  
	 * @param cell 指定单元格
	 * @return 1(单元格是合并区域的第一个单元)、2(合并区域内部的单元)、-1(不是合并区域内的单元)
	 * @throws ExcelNullParameterException 参数为null
	 */
	public static int isCellInMergedRegion(Sheet sheet,Cell cell) throws ExcelNullParameterException{
		if (hasMerged(sheet)) {
			if(cell == null){
				throw new ExcelNullParameterException();
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
	 * @return 1(单元格是合并区域的第一个单元)、2(合并区域内部的单元)、-1(不是合并区域内的单元)
	 * @throws ExcelIndexOutOfBoundsException 参数小于零
	 */
	public static int isCellInMergedRegion(Sheet sheet,int rowIndex,int columnIndex) throws ExcelIndexOutOfBoundsException{
		if (hasMerged(sheet)) {
			if(rowIndex <0 || columnIndex <0){
				throw new ExcelIndexOutOfBoundsException();
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
	 * 按类型返回公式计算值
	 * @param cellValue 
	 * @return 数值或文本类型返回具体值，其他返回""。
	 */
	private static String getFormulaValue(CellValue cellValue) {
        String value = "";
        switch (cellValue.getCellTypeEnum()) {
        case STRING:
            value = cellValue.getStringValue();
            break;
        case NUMERIC:
            value = String.valueOf(cellValue.getNumberValue());
            break;
        default:
            break;
        }
        return value;
    }
	
	/**
	 * 判断数据长度、起始下标和读取长度参数是否越界
	 * @param count 数据总长度,不能小于1
	 * @param startIndex 起始下标不能小于零或大于最大值
	 * @param length 读取的长度,不能小于0
	 * @return 如果参数越界抛出异常,否则返回要读取的最后一个下标值(如果读取长度大于总长度,则返回最大下标值)
	 * @throws ExcelIndexOutOfBoundsException 参数越界错误
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		if(count<1){ //数据长度小于1
			throw new ExcelIndexOutOfBoundsException();
		}
		if(length<0){//读取长度小于零
			throw new ExcelIndexOutOfBoundsException();
        }
        if(startIndex > count -1 || startIndex < 0){//开始下标大于最大的下标值或小于零
        	throw new ExcelIndexOutOfBoundsException();
        }
        //要读取的最后一个下标,如果下标越界，则读取至最后一个值
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
}
