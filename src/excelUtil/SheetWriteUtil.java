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
	
	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
	/**
	 * ʵ����sheet
	 */
	private Sheet sheet;
	
	/**
	 * ��sheet������������
	 */
	//private ArrayList<Row> allRowList = null;
	
	/**
	 * Excel������
	 */
	private Workbook workbook = null;
	
	//���кϲ�����ļ�ֵ�±�
	private int minRowIndexOfMergedRange = 0;
	private int maxRowIndexOfMergedRange = 0;
	private int minColumnIndexOfMergedRange = 0;
	private int maxColumnIndexOfMergedRange = 0;
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	
	
	
	

	/**
	 * ����дsheet�Ĺ���
	 * @param sheet ��Ҫд�����ݵ�sheet,��ExcelWriteUtil��ȡ
	 * @throws IllegalArgumentException sheetΪnull
	 */
	public SheetWriteUtil(Sheet sheet) throws IllegalArgumentException{
		if(sheet == null)
			throw new IllegalArgumentException("sheetΪnull");
		this.sheet = sheet;
		this.workbook = sheet.getWorkbook();
	}
	
	/**
	 * ����ĳ���Զ������п�
	 * @param columnIndex Ҫ�����п���±�
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
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
	
	/**
	 * ��ȡ��Ч��Cell��Ԫ(�Ǻϲ������ڲ��ĵ�Ԫ)
	 * @param rowIndex ���±�
	 * @param columnIndex ���±�
	 * @return ָ����Ԫ��,����null(��Ч��Ԫ��),�����׳��쳣
	 * @throws IndexOutOfBoundsException �±����С����
	 * @throws IllegalArgumentException columnIndex < 0 ���� �����ļ��ṩ���ֵ
	 */
	public Cell getValidCell(int rowIndex,int columnIndex) throws IndexOutOfBoundsException,IllegalArgumentException{
		try {
			Cell cell = null;
			if(hasMerged()){//����кϲ���Ԫ��
				int result = isCellInMergedRegion(rowIndex, columnIndex);
				if(result == 1){ //��Ԫ���Ǻϲ������һ��Ԫ
					cell = getCell(rowIndex, columnIndex);
				}else if(result == 2){ // ��Ԫ���Ǻϲ������ڲ��ĵ�Ԫ
					cell = null;
				}else { // ��Ԫ���Ǻϲ������ڵĵ�Ԫ
					cell = getCell(rowIndex, columnIndex);
				}
			}else{//û�кϲ�����
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
	 * ��ȡָ����Cell
	 * @param rowIndex ���±�
	 * @param columnIndex ���±�
	 * @return ����Cell,�����׳��쳣
	 * @throws IllegalArgumentException  columnIndex <0 ���� �����ļ��ṩ���ֵ
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
	 * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��  
	 * @param rowIndex ��Ԫ�����±�
	 * @param columnIndex ��Ԫ�����±�
	 * @return 1(��Ԫ���Ǻϲ������һ����Ԫ)��2(�ϲ������ڲ��ĵ�Ԫ)��-1(���Ǻϲ������ڵĵ�Ԫ)
	 * @throws IndexOutOfBoundsException ����С����
	 */
	public int isCellInMergedRegion(int rowIndex,int columnIndex) throws IndexOutOfBoundsException{
		if (hasMerged()) {
			if(rowIndex <0 || columnIndex <0){
				throw new IndexOutOfBoundsException("����С����");
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(rowIndex == firstRow && columnIndex == firstColumn){//��Ԫ��Ϊ�Ǻϲ�����ĵ�һ����Ԫ
					return 1;
				}else if (rowIndex >= firstRow && rowIndex <= lastRow) {//��Ԫ���ںϲ������ڲ�
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
						return 2;
					}
				}
			}
			return -1;//��Ԫ���ںϲ�������
		}else{
			return -1;
		}
	}
	
	/**  
	* �ж�sheetҳ���Ƿ��кϲ���Ԫ��   
	* @param sheet   
	* @return  �кϲ���Ԫ�񷵻�true���򷵻�false
	*/  
	public boolean hasMerged() {  
	     return sheet.getNumMergedRegions() > 0 ? true : false;  
	} 

}
