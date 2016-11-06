package excelUtil;
/** 
 * @author HPC
 * 
 */
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class SheetReadUtil {

	/**
	 * ��־����
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * ʵ����sheet
	 */
	private Sheet sheet;
	
	/**
	 * ��sheet������������
	 */
	private ArrayList<Row> allRowList = null;
	
	/**
	 * ��ѡ����(IntegerΪ��������±�,StringΪ��������)
	 */
	private HashMap<Integer, String> titles;

	/**
	 * @param sheet ��Ҫ������sheetʵ��
	 */
	public SheetReadUtil(Sheet sheet){
		this.sheet = sheet;
		allRowList = new ArrayList<>();
		titles = new HashMap<>();
	}
	
	/**
	 * ��allRowList�����һ��Row
	 * @param row
	 */
	public void addRow(Row row){
		allRowList.add(row);
	}

	/**
	 * ��ȡʵ����sheet
	 * @return
	 */
	public Sheet getSheet(){
		return sheet;
	}
	
	/**
	 * ����allRowList�б�
	 * @return
	 */
	public ArrayList<Row> getAllRowList() {
		return allRowList;
	}
	
	/**
	 * ��ȡallRowList�б���ָ����Row
	 * @param rowIndex Row�±�ֵ
	 * @return ����ָ��Row,�����׳��쳣
	 * @throws IndexOutOfBoundsException �±�ֵԽ���쳣 
	 */
	public Row getRowAt(int rowIndex) throws IndexOutOfBoundsException {
		try {
			return allRowList.get(rowIndex);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ͨ��ָ����������ֵ�����ݼ����ȡָ����RowList�б�
	 * @param rows Row�б����ݼ�
	 * @param title ��������
	 * @param value ��������ֵ
	 * @return ����ָ��RowList�б�
	 * @throws IllegalArgumentException ����ֵ������
	 */
	public ArrayList<Row> getRowListByArg(ArrayList<Row> rows,String title,String value) throws IllegalArgumentException{
		int columnIndex = getTitleColIndexByValue(title);
		ArrayList<Row> rowList = null;
		if(columnIndex != -1){//titleֵ����
			rowList = new ArrayList<>();
			for(Row row : rows){
				String v = getCellValue(row.getCell(columnIndex));
				if(v.equals(value)){
					rowList.add(row);
				}
			}
			return rowList;
		}else{
			throw new IllegalArgumentException("titleֵ������");
		}
	}
	
	/**
	 * ͨ��ָ����������ֵ�����ݼ���ȡָ����RowList�б�
	 * @param startRowIndex �������ݼ���allRowList�����ʼ�±�
	 * @param length �������ݼ��ĳ���
	 * @param title ��������
	 * @param value ��������ֵ
	 * @return ����ָ��RowList�б�
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 * @throws IllegalArgumentException ����ֵ������
	 */
	public ArrayList<Row> getRowListByArg(int startRowIndex,int length,String title,String value) throws
													IndexOutOfBoundsException,IllegalArgumentException{
		int count = allRowList.size();
		try {
			int endIndex = ExcelUtil.isIndexOutOfBounds(count, startRowIndex, length);
			ArrayList<Row> rows = new ArrayList<>();
			for(int i=startRowIndex;i<=endIndex;i++){
				rows.add(allRowList.get(i));
			}
			return getRowListByArg(rows, title, value);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}catch (IllegalArgumentException e) {
			throw e;
		}
	}
	
	/**
	 * ��ȡָ��Row������CellList
	 * @param row ָ��Row
	 * @return
	 */
	public ArrayList<Cell> getOneRowAllCells(Row row){
		int count = row.getLastCellNum();
		if(count == -1){
			return new ArrayList<Cell>();
		}
		return getOneRowCellList(row, 0, count);
	}
	
	/**
	 * ��ȡָ��Row��ָ����CellList
	 * @param row ָ��Row
	 * @param startColumnIndex ��ʼ�����±�ֵ
	 * @param length ���賤��ֵ
	 * @return
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 */
	public ArrayList<Cell> getOneRowCellList(Row row,int startColumnIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = row.getLastCellNum();
		if(count == -1){
			return cellList;
		}
		try {
			int endIndex = ExcelUtil.isIndexOutOfBounds(count, startColumnIndex, length);
			Cell c = null;
			for(int i=startColumnIndex;i<=endIndex;i++){
				c = row.getCell(i);
				cellList.add(c);
			}
			return cellList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ��ȡһ�е�����Cells
	 * @param columnIndex ָ�������±�ֵ
	 * @return ����CellList
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 */
	public ArrayList<Cell> getOneColumnAllCells(int columnIndex) throws IndexOutOfBoundsException{
		try {
			int length = allRowList.size();
			return getOneColumnCellList(columnIndex, 0, length);
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ��ȡһ����ָ����Cells
	 * @param cloumnIndex ָ�������±�ֵ
	 * @param startRowIndex ��ʼ�����±�ֵ
	 * @param length ���賤��
	 * @return 
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 */
	public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = allRowList.size();
		try {
			int endIndex = ExcelUtil.isIndexOutOfBounds(count, startRowIndex, length);
			Cell cell = null;
			Row row = null;
			for(int i=startRowIndex;i<=endIndex;i++){
				row = allRowList.get(i);
				if(row == null){
					cellList.add(null);
				}else{
					cell = row.getCell(cloumnIndex);
					cellList.add(cell);
				}
			}
			return cellList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * �趨sheet�ı����б�
	 * @param row ָ������
	 * @param startColumnIndex ��ʼ���±�ֵ
	 * @param length ���賤��
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 * @throws IllegalArgumentException ָ����û������
	 */
	public void setTitles(Row row,int startColumnIndex,int length) throws IndexOutOfBoundsException,
																IllegalArgumentException{
		int cellCount = row.getLastCellNum();
		if(cellCount == -1){
			throw new IllegalArgumentException("����û������");
		}
		try {
			ArrayList<Cell> cells = getOneRowCellList(row, startColumnIndex, length);
			Cell cell = null;
			for(int i=0;i<cells.size();i++){
				cell = cells.get(i);
				if(cell == null){
					titles.put(startColumnIndex+i,"");
				}else{
					titles.put(startColumnIndex+i, getCellValue(cell));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ��ȡ�����б�
	 * @return
	 */
	public HashMap<Integer, String> getTitles(){
		return titles;
	}
	
	/**
	 * ͨ������ֵ��ȡ�������±�
	 * @param titleName ָ������ֵ
	 * @return �������±�,�������򷵻�-1
	 */
	public int getTitleColIndexByValue(String titleName){
		for(Map.Entry<Integer, String> entry : titles.entrySet()){
			if(entry.getValue().equals(titleName)) 
				return entry.getKey();
		}
		return -1;
	}
	
	/**
	 * ��ȡָ��Cell������ֵ
	 * @param cell ָ��Cell
	 * @return ������������String���ͷ���
	 */
	@SuppressWarnings("deprecation")
	public String getCellValue(Cell cell){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			CellType cellType = cell.getCellTypeEnum();
			
			switch (cellType) {
			case STRING:
				string =  cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				short format = cell.getCellStyle().getDataFormat();
				/**
				 * �ж���������
				 * (2001.01.01) cellType:STRING format��0
				 * (yyyy-MM) cellType:NUM format:17 �����ڸ�ʽ
				 * (yyyy��MM��dd��) cellType:NUM format:31
				 * (yyyy��MM��) cellType:NUM format:57
				 * (MM-dd��MM��dd��) cellType:NUM format:58
				 * (yyyy-MM-dd) cellType:NUM format: 176 
				 * 177? ��ֵ�����ھ�����
				 * 179?
				 * 178?
				 */
				if(DateUtil.isCellDateFormatted(cell)){ //�������ڸ�ʽ
					SimpleDateFormat s = new SimpleDateFormat("yyyy-MM-dd");
					double d = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(d);
					string = s.format(date);
				}else if(format == 31 || format == 57 || format == 58 || format == 176){//�Զ������ڸ�ʽ
					SimpleDateFormat s = null;
					switch (format) {
					case 31:
						s = new SimpleDateFormat("yyyy��MM��dd��");
						break;
					case 57:
						s = new SimpleDateFormat("yyyy��MM��");
						break;
					case 58:
						s = new SimpleDateFormat("MM��dd��");
						break;
					case 176:
						s = new SimpleDateFormat("yyyy��MM��dd��");
						break;
					default:
						break;
					}
					double d = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(d);
					string = s.format(date);
				}else{
					double d = cell.getNumericCellValue();
					int in = (int)d;
					double last = d - in;
					if(last == 0) //doubleΪ����
						string = String.valueOf(in);
					else 
						string = String.valueOf(d);
				}
				break;
			case ERROR:
				string = String.valueOf(cell.getErrorCellValue());
				break;
			case FORMULA:
				string = cell.getCellFormula();
				break;
			case BOOLEAN:
				string = String.valueOf(cell.getBooleanCellValue());
				break;
			case BLANK:
				string = "";
				break;
			default:
				string = "";
			}
			return string.trim();
		} catch (Exception e) {
			logger.error("getCellValue()", e);
			return "";
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
	
	/**
	 * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��  
	 * @param cell ָ����Ԫ��
	 * @return ����Ǻϲ���Ԫ�񷵻�true���򷵻�false
	 * @throws IllegalArgumentException ����Ϊnull
	 */
	public boolean isMergedRegion(Cell cell) throws IllegalArgumentException{
		if (hasMerged()) {
			if(cell == null){
				throw new IllegalArgumentException("����Ϊnull");
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			int rowIndex = cell.getRowIndex();
			int colIndex = cell.getColumnIndex();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if (rowIndex >= firstRow && rowIndex <= lastRow) {
					if (colIndex >= firstColumn && colIndex <= lastColumn) {
						return true;
					}
				}
			}
			return false;
		}else{
			return false;
		}
	}
	
	/**
	 * ��ȡ�ϲ���Ԫ���ֵ 
	 * @param cell ָ����Ԫ��
	 * @return ��Ԫ���ֵ
	 * @throws IllegalArgumentException CellΪnull����cell���Ǻϲ���Ԫ��
	 * @throws IllegalStateException δ֪����(��Ӧ����)
	 */
	public String getCellValueOfMergedRegion(Cell cell) throws IllegalArgumentException,IllegalStateException{  
		try {
			if (isMergedRegion(cell)) {
				int sheetMergeCount = sheet.getNumMergedRegions();
				int rowIndex = cell.getRowIndex();
				int columnIndex = cell.getColumnIndex();
				for (int i = 0; i < sheetMergeCount; i++) {
					CellRangeAddress ca = sheet.getMergedRegion(i);
					int firstColumn = ca.getFirstColumn();
					int lastColumn = ca.getLastColumn();
					int firstRow = ca.getFirstRow();
					int lastRow = ca.getLastRow();
					
					if (rowIndex >= firstRow && rowIndex <= lastRow) {
						if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
							Row fRow = sheet.getRow(firstRow);
							Cell fCell = fRow.getCell(firstColumn);
							return getCellValue(fCell);
						}
					}
				}
				throw new IllegalStateException("δ֪����");
			}
			else{
				throw new IllegalArgumentException("ָ����Ԫ���Ǻϲ���Ԫ��");
			}   
		} catch (IllegalArgumentException e) {
			throw e;
		}
	}    
	
}
