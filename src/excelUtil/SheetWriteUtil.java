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
 *дsheet��Ĺ���
 */
public class SheetWriteUtil {
	
	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
	/**
	 * ʵ����sheet
	 */
	protected Sheet sheet;
	
	/**
	 * ��sheet������������
	 */
	//private ArrayList<Row> rowList = null;
	
	/**
	 * Excel������
	 */
	protected Workbook workbook = null;
	
	//���кϲ�����ļ�ֵ�±�
	protected int minRowIndexOfMergedRange = 0;
	protected int maxRowIndexOfMergedRange = 0;
	protected int minColumnIndexOfMergedRange = 0;
	protected int maxColumnIndexOfMergedRange = 0;
	
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
	
	
	public Sheet getSheet(){
		return this.sheet;
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
	 * ��ȡ��Ч��
	 * @param rowIndex ���±�
	 * @return Rowʵ��
	 */
	public Row getValidRow(int rowIndex){
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			row = sheet.createRow(rowIndex);
		return row;
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
	 * �������������Ԫ����
	 * @param cell ��Ԫ��(��Ϊnull����ִ�и÷���)
	 * @param value ����ֵ (����ֵ�������(yyyy-MM-dd,yyyy-MM,MM-dd))
	 * @param cellType ��Ԫ����������
	 * @param cellStyle ��Ԫ����ʽ(��Ϊnull����cellTypeΪ���ڸ�ʽ���贫���µ�cellStyleʵ��)
	 * @throws (-----��ϸ��Ϣ������message��-------)
	 * @throws IllegalArgumentException  ����ֵ��ʽ����
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
			if (last == 0) // doubleΪ����
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
			//Value��ʽΪ((yyyy-MM-dd,yyyy-MM,MM-dd))
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
	 * ����ĳ���Զ������п�
	 * @param columnIndex Ҫ�����п���±�
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
	}
	
	/**
	 * ��ȡָ����Cell
	 * @param rowIndex ���±�
	 * @param columnIndex ���±�
	 * @return ����Cell,�����׳��쳣
	 * @throws IllegalArgumentException  columnIndex <0 ���� �����ļ��ṩ���ֵ
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
