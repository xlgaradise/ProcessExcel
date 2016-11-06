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
	 * ��־����
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * ʵ����sheet
	 */
	private Sheet sheet;
	
	/**
	 * sheet����
	 */
	private String sheetName = "";

	/**
	 * ��sheet������������
	 */
	private ArrayList<Row> allRowList = null;
	
	/**
	 * sheet����Excel�ļ���ʽ,Ĭ��Ϊ.xls
	 */
	private String extension = ".xls";
	
	/**
	 * Excel������
	 */
	private Workbook wb = null;
	
	//���кϲ�����ļ�ֵ�±�
	private int minRowIndexOfMergedRange = 0;
	private int maxRowIndexOfMergedRange = 0;
	private int minColumnIndexOfMergedRange = 0;
	private int maxColumnIndexOfMergedRange = 0;
	
	/**
	 * @param sheetName sheet����
	 */
	public SheetWriteUtil(String sheetName){
		this.sheetName = sheetName;
		wb = new HSSFWorkbook();
	}
	
	/**
	 * @param sheetName sheet����
	 * @param extension sheet����Excel�ļ��ĺ�׺(.xls��.xlsx)
	 * @throws IllegalArgumentException  ��׺������
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
			throw new IllegalArgumentException("��׺������");
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
	 * ����ͨ�õ�Ԫ����ʽ(ˮƽ���о��С���ֱ���С��Զ�����)
	 * @return
	 */
	public CellStyle getCommonCellStyle_alignCenter(){
		CellStyle cellStyle = wb.createCellStyle();
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
		CellStyle cellStyle = wb.createCellStyle();
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
		CellStyle cellStyle = wb.createCellStyle();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
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
		Font font = wb.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short)12);
		font.setBold(true);
		return font;
	}
	
	/**
	 * ��Ӻϲ���Ԫ��
	 * @param startRowIndex ��ʼ���±�
	 * @param endRowIndex �������±�
	 * @param startColumnIndex ��ʼ���±�
	 * @param endColumnIndex �������±�
	 * @throws IllegalArgumentException  endIndex С�� startIndex
	 */
	public void addMergedRegion(int startRowIndex,int endRowIndex,int startColumnIndex,int endColumnIndex)
												throws IllegalArgumentException{
		if(startRowIndex > endRowIndex || startColumnIndex > endColumnIndex){
			throw new IllegalArgumentException("endIndex С�� startIndex");
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
