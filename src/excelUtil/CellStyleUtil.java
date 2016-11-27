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
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *��Ԫ����ʽ����
 */
public class CellStyleUtil {
	
	/*%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	
	/*%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%*/

	
	
	/**
	 * ������Ԫ����ʽ����
	 * @param workbook
	 * @throws IllegalArgumentException workbookΪnull
	 */
	public CellStyleUtil(Workbook workbook) throws IllegalArgumentException{
		if(workbook == null){
			throw new IllegalArgumentException();
		}
		this.workbook = workbook;
	}
	
	/**
	 * ���ص�Ԫ����ʽʵ��
	 * @return
	 */
	public CellStyle getCellStyle(){
		CellStyle cellStyle = workbook.createCellStyle();
		return cellStyle;
	}
	
	/**
	 * ����ͨ�õ�Ԫ����ʽ(ˮƽ���о��С���ֱ���С��Զ�����)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignCenter(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION); //ˮƽ���о���
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //��ֱ����
		cellStyle.setWrapText(true);//�Զ�����
		return cellStyle;
	}
	
	/**
	 * ����ͨ�õ�Ԫ����ʽ(ˮƽ���󡢴�ֱ���С��Զ�����)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignLeft(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.LEFT); //ˮƽ����
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //��ֱ����
		cellStyle.setWrapText(true);//�Զ�����
		return cellStyle;
	}
	
	/**
	 * ����ͨ�õ�Ԫ����ʽ(ˮƽ���ҡ���ֱ���С��Զ�����)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignRight(){
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setAlignment(HorizontalAlignment.RIGHT); //ˮƽ����
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER); //��ֱ����
		cellStyle.setWrapText(true);//�Զ�����
		return cellStyle;
	}
	
	/**
	 * ����ͨ������������ʽ(���塢��С10�����Ӵ�)
	 * @return
	 */
	public Font getCommonFont_content(){
		Font font = workbook.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short)10);
		font.setBold(false);
		return font;
	}
	
	/**
	 * ����ͨ������������ʽ(���塢��С10���Ӵ�)
	 * @return
	 */
	public Font getCommonFont_content_bold(){
		Font font = workbook.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		return font;
	}
	
	/**
	 * ����ͨ������������ʽ(���塢��С10�����Ӵ֡���ɫ)
	 * @return
	 */
	public Font getCommonFont_content_red(){
		Font font = workbook.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		font.setColor(Font.COLOR_RED);
		return font;
	}
	
	/**
	 * ����ͨ�ñ���������ʽ(���塢��С12���Ӵ�)
	 * @return
	 */
	public Font getCommonFont_title(){
		Font font = workbook.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short)10);
		font.setBold(true);
		return font;
	}
	
	/**
	 * ���õ�Ԫ���ǰ��ɫ��
	 * @param cellStyle ��Ҫ���õĵ�Ԫ��
	 * @param color ��ɫֵ
	 */
	public void setForegroundColor(CellStyle cellStyle,IndexedColors color){
		cellStyle.setFillForegroundColor(color.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}
	
}
