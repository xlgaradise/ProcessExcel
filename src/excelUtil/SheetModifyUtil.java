
package excelUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNullParameterException;

/**
*@auchor HPC
*@encoding GBK
*/

public class SheetModifyUtil {
	
	// %%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%

	/**
	 * ʵ����sheet
	 */
	protected Sheet sheet;

	/**
	 * Excel������
	 */
	protected Workbook workbook = null;

	// %%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%

	/**
	 * ����дsheet�Ĺ���
	 * 
	 * @param sheet ��Ҫд�����ݵ�sheet,��ExcelModifyUtil��ȡ
	 * @throws ExcelNullParameterException sheetΪnull
	 */
	public SheetModifyUtil(Sheet sheet) throws ExcelNullParameterException {
		if (sheet == null)
			throw new ExcelNullParameterException();
		this.sheet = sheet;
		this.workbook = sheet.getWorkbook();
	}

	public Sheet getSheet() {
		return this.sheet;
	}

	
	/**
	 * ɾ��ĳһ��
	 * @param row
	 */
	public void removeRow(Row row){
		this.sheet.removeRow(row);
	}
	
	
	/**
	 * ɾ��ĳһ��
	 * @param rowIndex
	 * @throws ExcelIllegalArgumentException �±�ֵ������Ч��Χ
	 */
	public void removeRowAt(int rowIndex) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(rowIndex < 0 || rowIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = this.sheet.getRow(rowIndex);
		removeRow(row);
	}
	
	
	/**
	 * ɾ��ĳЩ��
	 * @param startIndex (base 0)
	 * @param length 
	 * @throws ExcelIllegalArgumentException �±�ֵ������Ч��Χ
	 */
	public void removeRowsIn(int startIndex,int length) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(startIndex < 0 || startIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}else if((startIndex + length - 1) > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = null;
		for(int i=startIndex+length-1;i>=startIndex;i--){
			row = sheet.getRow(i);
			removeRow(row);
		}
	}
	
	/**
	 * ɾ��ָ���±���������
	 * @param startIndex
	 * @throws ExcelIllegalArgumentException �±�ֵ������Ч��Χ
	 */
	public void removeRowsFrom(int startIndex) throws ExcelIllegalArgumentException{
		int endIndex = sheet.getLastRowNum();
		if(startIndex < 0 || startIndex > endIndex){
			throw new ExcelIllegalArgumentException();
		}
		Row row = null;
		for(int i=endIndex;i>=startIndex;i--){
			row = sheet.getRow(i);
			removeRow(row);
		}
	}
	
	/**
	 * ��ȡĳһ��
	 * @param rowIndex
	 * @return Rowʵ����null
	 */
	public Row getRowAt(int rowIndex){
		return sheet.getRow(rowIndex);
	}
	
	/**
	 * ��ָ���±괴���µ�һ��
	 * @param rowIndex
	 * @return Rowʵ��
	 */
	public Row createNewRow(int rowIndex){
		return sheet.createRow(rowIndex);
	}
	
	/**
	 * ��ȡ��Ч��Cell��Ԫ(�Ǻϲ������ڲ��ĵ�Ԫ)
	 * @param rowIndex ���±�
	 * @param columnIndex ���±�
	 * @return ָ����Ԫ��,����null(��Ч��Ԫ��)
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
	 * @throws ExcelIllegalArgumentException  ����ֵ��ƥ���Ӧ��ʽ
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
			cellStyle = cellStyleUtil.getCommonCellStyle();
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
	 * ����ĳ���Զ������п�
	 * @param columnIndex Ҫ�����п���±�
	 */
	public void setAutoSizeColumn(int columnIndex){
		sheet.autoSizeColumn(columnIndex);
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
}
