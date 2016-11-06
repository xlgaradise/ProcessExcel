package excelUtil;

import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *@auchor HPC
 *
 */
public class SheetWriteUtil {
	/**
	 * 日志工具
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * 实例的sheet
	 */
	private Sheet sheet;
	
	/**
	 * sheet名字
	 */
	private String sheetName = "";

	/**
	 * 该sheet的所有行数据
	 */
	private ArrayList<Row> allRowList = null;
	
	/**
	 * sheet所属Excel文件格式,默认为.xls
	 */
	private String extension = ".xls";
	
	/**
	 * Excel工作薄
	 */
	private Workbook wb = null;
	
	//所有合并区域的极值下标
	private int minRowIndexOfMergedRange = 0;
	private int maxRowIndexOfMergedRange = 0;
	private int minColumnIndexOfMergedRange = 0;
	private int maxColumnIndexOfMergedRange = 0;
	
	/**
	 * @param sheetName sheet名称
	 */
	public SheetWriteUtil(String sheetName){
		this.sheetName = sheetName;
		wb = new HSSFWorkbook();
	}
	
	/**
	 * @param sheetName sheet名称
	 * @param extension sheet所属Excel文件的后缀(.xls或.xlsx)
	 * @throws IllegalArgumentException  后缀名错误
	 */
	public SheetWriteUtil(String sheetName,String extension) throws IllegalArgumentException{
		if(extension.equals(".xls")){
			this.sheetName = sheetName;
			this.extension = extension;
			this.wb = new HSSFWorkbook();
			this.sheet = wb.createSheet();
			
		}else if (extension.equals(".xlsx")) {
			this.sheetName = sheetName;
			this.extension = extension;
			this.wb = new XSSFWorkbook();
			this.sheet = wb.createSheet();
		}else {
			throw new IllegalArgumentException("后缀名错误");
		}
	}
	
	public Sheet getSheet() {
		return sheet;
	}

	public String getSheetName() {
		return sheetName;
	}

	public String getExtension() {
		return extension;
	}
	
	
	/**
	 * 返回通用单元格样式(水平跨列居中、垂直居中、自动换行)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignCenter(){
		CellStyle cellStyle = wb.createCellStyle();
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
		CellStyle cellStyle = wb.createCellStyle();
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
		CellStyle cellStyle = wb.createCellStyle();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
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
	
	public void addCellValue(int rowIndex,Cell cell){
		
	}

}
