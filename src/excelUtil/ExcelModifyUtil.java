/**
*@author HPC
*/
package excelUtil;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import excelUtil.CellTypeUtil.TypeEnum;

public class ExcelModifyUtil {
	
	/*%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%*/
	
	protected Workbook workbook = null;
	protected String path = "";
	
	/*%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%*/
	
	
	/**
	 * ����Excel�ļ��޸Ĺ���
	 * @param excelReadUtil
	 * @throws IllegalArgumentException ����Ϊnull
	 */
	public ExcelModifyUtil(ExcelReadUtil excelReadUtil) throws IllegalArgumentException{
		if(excelReadUtil == null){
			throw new IllegalArgumentException();
		}
		this.workbook = excelReadUtil.getWorkBook();
		this.path = excelReadUtil.getFilePath();
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
			DataFormat	dataFormat_num = this.workbook.createDataFormat();
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
			DataFormat dataFormat = this.workbook.createDataFormat();
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
	 * ��workbook����д��Excel�ļ�
	 * @throws FileNotFoundException �ļ��ѱ���һ�������
	 * @throws SecurityException �ܾ����ļ�����д�����
	 * @throws IOException �ļ�д����߹رճ���
	 */
	public void writeToExcel() throws FileNotFoundException,SecurityException,IOException{
		try {
			FileOutputStream outputStream = new FileOutputStream(this.path);
			workbook.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (FileNotFoundException e) {
			throw e;
		} catch (SecurityException e) {
			throw e;
		} catch (IOException e) {
			throw e;
		}
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
		
		CellStyleUtil cellStyleUtil = new CellStyleUtil(workbook);
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
		CellStyleUtil cellStyleUtil = new CellStyleUtil(workbook);
		CellStyle cellStyle = cellStyleUtil.getCommonCellStyle_alignCenter();
		cellStyle.setWrapText(false);
		cellStyleUtil.setForegroundColor(cellStyle, color);
		cell.setCellStyle(cellStyle);
	}
	
	/**
	 * �ر�д�빤��
	 * @throws IOException
	 */
	public void close() throws IOException{
		try {
			this.workbook.close();
		} catch (IOException e) {
			throw e;
		}
	}
	
}
