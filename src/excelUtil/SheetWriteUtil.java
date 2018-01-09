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
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNullParameterException;


/**
 *дsheet��Ĺ���
 */
public class SheetWriteUtil {
	
	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
	/**
	 * ʵ����sheet
	 */
	protected SXSSFSheet sheet;
	
	/**
	 * ��sheet������������
	 */
	//private ArrayList<Row> rowList = null;
	
	/**
	 * Excel������
	 */
	protected SXSSFWorkbook workbook = null;
	
	//���кϲ�����ļ�ֵ�±�
	protected int minRowIndexOfMergedRange = 0;
	protected int maxRowIndexOfMergedRange = 0;
	protected int minColumnIndexOfMergedRange = 0;
	protected int maxColumnIndexOfMergedRange = 0;
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	
	
	
	

	/**
	 * ����дsheet�Ĺ���
	 * @param sheet ��Ҫд�����ݵ�sheet,��ExcelWriteUtil��ȡ
	 * @throws ExcelNullParameterException sheetΪnull
	 */
	public SheetWriteUtil(SXSSFSheet sheet) throws ExcelNullParameterException{
		if(sheet == null)
			throw new ExcelNullParameterException();
		this.sheet = sheet;
		this.workbook = sheet.getWorkbook();
		this.sheet.trackAllColumnsForAutoSizing();
		
	}
	
	
	public SXSSFSheet getSheet(){
		return this.sheet;
	}

	public Workbook getWorkBook(){
		return this.workbook;
	}
	
	
	/**
	 * ��Ӻϲ���Ԫ��
	 * @param startRowIndex ��ʼ���±�
	 * @param endRowIndex �������±�
	 * @param startColumnIndex ��ʼ���±�
	 * @param endColumnIndex �������±�
	 * @throws ExcelIllegalArgumentException  endIndex С�� startIndex
	 */
	public void addMergedRegion(int startRowIndex,int endRowIndex,int startColumnIndex,int endColumnIndex)
												throws ExcelIllegalArgumentException{
		if(startRowIndex > endRowIndex || startColumnIndex > endColumnIndex){
			throw new ExcelIllegalArgumentException();
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
		if(row == null){
			row = sheet.createRow(rowIndex);
			row.setHeightInPoints(21);
		}
		return row;
	}
	
	/**
	 * ��ȡ��Ч��Cell��Ԫ(�Ǻϲ������ڲ��ĵ�Ԫ)
	 * @param rowIndex ���±�
	 * @param columnIndex ���±�
	 * @return ָ����Ԫ��,����null(��Ч��Ԫ��),�����׳��쳣
	 * @throws ExcelIndexOutOfBoundsException �±����С����
	 * @throws IllegalArgumentException columnIndex < 0 ���� �����ļ��ṩ���ֵ
	 */
	public Cell getValidCell(int rowIndex,int columnIndex) throws ExcelIndexOutOfBoundsException,IllegalArgumentException{
		
		Cell cell = null;
		if (SheetReadUtil.hasMerged(sheet)) {// ����кϲ���Ԫ��
			int result = SheetReadUtil.isCellInMergedRegion(sheet,rowIndex, columnIndex);
			if (result == 1) { // ��Ԫ���Ǻϲ������һ��Ԫ
				cell = getCell(rowIndex, columnIndex);
			} else if (result == 2) { // ��Ԫ���Ǻϲ������ڲ��ĵ�Ԫ
				cell = null;
			} else { // ��Ԫ���Ǻϲ������ڵĵ�Ԫ
				cell = getCell(rowIndex, columnIndex);
			}
		} else {// û�кϲ�����
			cell = getCell(rowIndex, columnIndex);
		}
		return cell;
		
	}
	
	/**
	 * �������������Ԫ����
	 * @param cell ��Ԫ��(��Ϊnull����ִ�и÷���)
	 * @param value ����ֵ (����ֵ�������(yyyy-MM-dd,yyyy-MM,MM-dd))
	 * @param cellType ��Ԫ����������
	 * @param cellStyle ��Ԫ����ʽ(��Ϊnull����cellTypeΪ���ڸ�ʽ���贫���µ�cellStyleʵ��)
	 * @throws ExcelIllegalArgumentException  ���ݸ�ʽ��Ӧ��ֵ����
	 */
	public void addValueToCell(Cell cell,String value,
			TypeEnum cellType,CellStyle cellStyle) throws ExcelIllegalArgumentException{
		
		if(cell == null){
			return;
		}
		if(cellStyle == null){
			CellStyleUtil cellStyleUtil = null;
			try {
				cellStyleUtil = new CellStyleUtil(workbook);
			} catch (ExcelNullParameterException e) {
			}
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
			}catch (NullPointerException |NumberFormatException e) {
				throw new ExcelIllegalArgumentException();
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
				throw new ExcelIllegalArgumentException();
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
			if(length >= 8 && length <= 10){
				simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
			 	formatNum = dataFormat.getFormat("yyyy-MM-dd");
			}else if (length == 6 || length == 7) {
				simpleDateFormat = new SimpleDateFormat("yyyy-MM");
				formatNum = dataFormat.getFormat("yyyy-MM");
			}else if(length >= 3 && length <= 5){
				simpleDateFormat = new SimpleDateFormat("MM-dd");
				formatNum = dataFormat.getFormat("MM-dd");
			}else{
				throw new ExcelIllegalArgumentException();
			}
			try {
				dateStr = simpleDateFormat.parse(value);
			} catch (ParseException e) {
				throw new ExcelIllegalArgumentException();
			}
			cellStyle.setDataFormat(formatNum);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(dateStr);
			break;
		case ERROR:
			cellStyle.setFillForegroundColor(IndexedColors.RED.index);
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			cell.setCellStyle(cellStyle);
			//cell.setCellErrorValue(Byte.parseByte(value));
			cell.setCellValue(value);
			break;
		case FORMULA:
			cell.setCellStyle(cellStyle);
			//cell.setCellFormula(value);
			dd = 0;
			try{
				dd = Double.parseDouble(value);
				in = (int) dd;
				last = dd - in;
				if (last == 0) // doubleΪ����
					cell.setCellValue(in);
				else
					cell.setCellValue(dd);
			}catch (NullPointerException |NumberFormatException e) {
				cell.setCellValue(value);
			}
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
	 * ������������
	 * @param colSplit �������������(base 1)
	 * @param rowSplit �ϱ�����������(base 1)
	 */
	public void setFreezePane(int colSplit, int rowSplit){
		sheet.createFreezePane(colSplit,rowSplit);
	}
	
	/**
	 * ����ָ���е�ǰ��ɫ
	 * @param row Rowʵ��(���RowΪnull����û��Cell��ִ�з���)
	 * @param color ��ɫ
	 */
	public void setForegroundColor(Row row,IndexedColors color){
		if(row == null){
			return;
		}
		int count = row.getLastCellNum();
		if(count == -1){
			return;
		}
		
		CellStyleUtil cellStyleUtil = null;
		try {
			cellStyleUtil = new CellStyleUtil(workbook);
		} catch (ExcelNullParameterException e) {
		}
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		for(int i=0;i<count;i++){
			Cell cell = row.getCell(i);
			if(cell == null){
				cell = row.createCell(i);
			}
			cellStyleUtil.setForegroundColor(cellStyle, color);
			cell.setCellStyle(cellStyle);
		}
	}
	
	/**
	 * ����ָ����Ԫ���ǰ��ɫ
	 * @param cell ָ����Ԫ��(���CellΪnull��ִ�з���)
	 * @param color ָ����ɫ
	 */
	public void setForegroundColor(Cell cell,IndexedColors color){
		if(cell == null)
			return;
		CellStyleUtil cellStyleUtil = null;
		try {
			cellStyleUtil = new CellStyleUtil(workbook);
		} catch (ExcelNullParameterException e) {
		}
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		cellStyleUtil.setForegroundColor(cellStyle, color);
		cell.setCellStyle(cellStyle);
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
		if(row == null){
			row = sheet.createRow(rowIndex);
			row.setHeightInPoints(21);
		}
		Cell cell = row.getCell(columnIndex);
		if(cell == null)
			try {
				cell = row.createCell(columnIndex);
			} catch (IllegalArgumentException e) {
				throw e;
			}
		return cell;
	}

}
