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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class SheetReadUtil {

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
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
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * ������sheet�Ĺ���
	 * @param sheet ��Ҫ������sheetʵ��,��ExcelReadUtil��ȡ
	 */
	public SheetReadUtil(Sheet sheet){
		this.sheet = sheet;
		allRowList = new ArrayList<>();
		titles = new HashMap<>();
	}
	

	/**
	 * ��ȡʵ����sheet
	 * @return ���ظ�SheetReadUtil������sheet
	 */
	public Sheet getSheet(){
		return sheet;
	}
	
	/**
	 * ��ȡSheet�����е���
	 */
	public void readAllRows(){
		int rowsCount = sheet.getLastRowNum() + 1;
		readRows(0, rowsCount);
	}

	/**
	 * ��ȡsheet��ָ������
	 * @param startIndex ��ʼ�����±�
	 * @param length ��ȡ����
	 * @throws IndexOutOfBoundsException ����Խ�����
	 */
	public void readRows(int startIndex,int length) throws IndexOutOfBoundsException{
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
			for(int i = startIndex;i<=endIndex;i++){
				if(endIndex == 0){//ֻ��sheet�ĵ�һ��
					Row r = sheet.getRow(0);
					if(r != null)
						this.allRowList.add(r);
				}else{
					this.allRowList.add(sheet.getRow(i));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
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
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
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
	 * @return ����cellList
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
	 * @return ����cellList�����׳�����
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 */
	public ArrayList<Cell> getOneRowCellList(Row row,int startColumnIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = row.getLastCellNum();
		if(count == -1){
			return cellList;
		}
		try {
			int endIndex = isIndexOutOfBounds(count, startColumnIndex, length);
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
	 * @return ����cellList�����׳�����
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
	 * @return ����cellList�б�����׳�����
	 * @throws IndexOutOfBoundsException �±�ֵԽ��
	 */
	public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = allRowList.size();
		try {
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
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
	 * @return ���ر���
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
		} catch (Exception e) {//δ֪����
			return "";
		}
	}
	
	/**
	 * ��ȡ�ϲ���Ԫ���ֵ 
	 * @param cell ָ����Ԫ��
	 * @return ��Ԫ���ֵ
	 * @throws IllegalArgumentException ����cellΪnull,����cell���Ǻϲ���Ԫ��
	 */
	public String getCellValueOfMergedRegion(Cell cell) throws IllegalArgumentException{  
		try {
			int result = isCellInMergedRegion(cell);
			if (result == 2) { //��Ԫ���ںϲ������ڲ�
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
						if (columnIndex >= firstColumn && columnIndex <= lastColumn) { //ȷ������λ��
							Row fRow = sheet.getRow(firstRow);
							Cell fCell = fRow.getCell(firstColumn);
							return getCellValue(fCell); //���ظ�����ĵ�һ����Ԫ���ֵ
						}
					}
				}
				return "";
			}else if (result == 1) { //��Ԫ��Ϊ�ϲ�����ĵ�һ����Ԫ
				return getCellValue(cell); //ֱ�ӷ��ظõ�Ԫ���ֵ
			}else{
				throw new IllegalArgumentException("ָ����Ԫ���Ǻϲ���Ԫ��");
			}   
		} catch (IllegalArgumentException e) {
			throw e;
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
	 * @return 1(��Ԫ���Ǻϲ������һ����Ԫ)��2(�ϲ������ڲ��ĵ�Ԫ)��-1(���Ǻϲ������ڵĵ�Ԫ)
	 * @throws IllegalArgumentException ����Ϊnull
	 */
	public int isCellInMergedRegion(Cell cell) throws IllegalArgumentException{
		if (hasMerged()) {
			if(cell == null){
				throw new IllegalArgumentException("����Ϊnull");
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			int rowIndex = cell.getRowIndex();
			int columnIndex = cell.getColumnIndex();
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
			return -1;
		}else{
			return -1;
		}
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
	 * �ж����ݳ��ȡ���ʼ�±�Ͷ�ȡ���Ȳ����Ƿ�Խ��
	 * @param count �����ܳ���,����С��1
	 * @param startIndex ��ʼ�±겻��С�����������ֵ
	 * @param length ��ȡ�ĳ���,����С��0
	 * @return �������Խ���׳��쳣,���򷵻�Ҫ��ȡ�����һ���±�ֵ(����±�Խ��,�򷵻�����±�ֵ)
	 * @throws IndexOutOfBoundsException ��������
	 */
	private int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("���ݳ���С��1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("��ȡ����С����");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("��ʼ�±���������±�ֵ��С����");
        }
        //Ҫ��ȡ�����һ���±�,����±�Խ�磬���ȡ�����һ��ֵ
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
	
}
