package excelUtil;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 *@auchor HPC
 *
 */
public class SheetWriteUtil {
	
	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	
	/**
	 * 实例的sheet
	 */
	private Sheet sheet;
	
	/**
	 * 该sheet的所有行数据
	 */
	//private ArrayList<Row> allRowList = null;
	
	/**
	 * Excel工作薄
	 */
	private Workbook workbook = null;
	
	//所有合并区域的极值下标
	private int minRowIndexOfMergedRange = 0;
	private int maxRowIndexOfMergedRange = 0;
	private int minColumnIndexOfMergedRange = 0;
	private int maxColumnIndexOfMergedRange = 0;
	
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
	
	/**
	 * 设置某列自动调整列宽
	 * @param columnIndex 要调整列宽的下标
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
	}
	
	/**
	 * 返回通用单元格样式(水平跨列居中、垂直居中、自动换行)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignCenter(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION); //水平跨列居中
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
		cellStyle.setWrapText(true);//自动换行
		return cellStyle;
	}
	
	/**
	 * 返回通用单元格样式(水平靠左、垂直居中、自动换行)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignLeft(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.LEFT); //水平靠左
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
		cellStyle.setWrapText(true);//自动换行
		return cellStyle;
	}
	
	/**
	 * 返回通用单元格样式(水平靠右、垂直居中、自动换行)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignRight(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.RIGHT); //水平靠右
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
		cellStyle.setWrapText(true);//自动换行
		return cellStyle;
	}
	
	/**
	 * 返回通用内容字体样式(宋体、大小10、不加粗)
	 * @return
	 */
	public Font getCommonFont_content(){
		Font font = workbook.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short)10);
		font.setBold(false);
		return font;
	}
	
	/**
	 * 返回通用内容字体样式(宋体、大小10、加粗)
	 * @return
	 */
	public Font getCommonFont_content_bold(){
		Font font = workbook.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		return font;
	}
	
	/**
	 * 返回通用内容字体样式(宋体、大小10、不加粗、红色)
	 * @return
	 */
	public Font getCommonFont_content_red(){
		Font font = workbook.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		font.setColor(Font.COLOR_RED);
		return font;
	}
	
	/**
	 * 返回通用标题字体样式(黑体、大小12、加粗)
	 * @return
	 */
	public Font getCommonFont_title(){
		Font font = workbook.createFont();
		font.setFontName("黑体");
		font.setFontHeightInPoints((short)12);
		font.setBold(true);
		return font;
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
	 * 获取指定的Cell
	 * @param rowIndex 行下标
	 * @param columnIndex 列下标
	 * @return 返回Cell,或者抛出异常
	 * @throws IllegalArgumentException  columnIndex <0 或者 大于文件提供最大值
	 */
	private Cell getCell(int rowIndex,int columnIndex) throws IllegalArgumentException{
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
