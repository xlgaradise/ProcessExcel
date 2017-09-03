/**
*@auchor HPC
*
*/
package excelUtil;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFont;

import exception.ExcelNullParameterException;

/**
 *单元格样式工具
 */
public class CellStyleUtil {
	
	/*%%%%%%%%-------字段部分 开始----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	
	/*%%%%%%%%-------字段部分 结束----------%%%%%%%%%*/

	
	
	/**
	 * 创建单元格样式工具
	 * @param workbook
	 * @throws ExcelNullParameterException workbook为null
	 */
	public CellStyleUtil(Workbook workbook) throws ExcelNullParameterException{
		if(workbook == null){
			throw new ExcelNullParameterException();
		}
		this.workbook = workbook;
	}
	
	/**
	 * 返回单元格样式实例
	 * @return
	 */
	public CellStyle getCellStyle(){
		CellStyle cellStyle = workbook.createCellStyle();
		return cellStyle;
	}
	
	/**
	 * 返回通用单元格样式(水平跨列居中、自动换行、宋体、大小10、不加粗)
	 * @return
	 */
	public CellStyle getCommonCellStyle(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION); //水平跨列居中
		cellStyle.setWrapText(true);//自动换行
		
		XSSFFont contenetFont = (XSSFFont) getCommonFont_content();
		cellStyle.setFont(contenetFont);
		return cellStyle;
	}
	
	/**
	 * 返回通用单元格样式(水平跨列居中、垂直居中、自动换行)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignCenter(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION); //水平跨列居中
		//cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
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
		//cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
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
		//cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //垂直居中
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
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		return font;
	}
	
	/**
	 * 设置单元格的前景色、
	 * @param cellStyle 需要设置的单元格
	 * @param color 颜色值
	 */
	public void setForegroundColor(CellStyle cellStyle,IndexedColors color){
		cellStyle.setFillForegroundColor(color.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}
	
}
