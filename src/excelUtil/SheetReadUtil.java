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
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import excelUtil.CellTypeUtil.TypeEnum;

/**
 *��sheet��Ĺ���,�ɻ�ȡ�С��С���Ԫ���
 */
public class SheetReadUtil {

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
	/**
	 * ʵ����sheet
	 */
	protected Sheet sheet;
	
	/**
	 * ��sheet��ָ��������
	 */
	protected ArrayList<Row> rowList = null;
	
	/**
	 * ��ѡ����(IntegerΪ��������±�,StringΪ��������)
	 */
	protected HashMap<Integer, String> titles;
	
	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * ������sheet�Ĺ���
	 * @param sheet ��Ҫ������sheetʵ��,��ExcelReadUtil��ȡ
	 */
	public SheetReadUtil(Sheet sheet){
		this.sheet = sheet;
		rowList = new ArrayList<>();
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
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		readRows(0, rowsCount);
	}

	/**
	 * ��ȡsheet��ָ������
	 * @param startIndex ��ʼ�����±�
	 * @param length ��ȡ����
	 * @throws IndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 */
	public void readRows(int startIndex,int length) throws IndexOutOfBoundsException{
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
			for(int i = startIndex;i<=endIndex;i++){
				if(endIndex == 0){//ֻ��sheet�ĵ�һ��
					Row r = sheet.getRow(0);
					if(r != null)
						this.rowList.add(r);
				}else{
					this.rowList.add(sheet.getRow(i));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * ����rowList�б�
	 * @return
	 */
	public ArrayList<Row> getRowList() {
		return rowList;
	}
	
	/**
	 * ��ȡrowList�б���ָ����Row
	 * @param rowIndex Row�±�ֵ(��rowList�е��±������Excel�е����±�)
	 * @return ����ָ��Row,�����׳��쳣
	 * @throws IndexOutOfBoundsException �±�ֵԽ���쳣 
	 */
	public Row getRowAt(int rowIndex) throws IndexOutOfBoundsException {
		try {
			return rowList.get(rowIndex);
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
	 * @throws IndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 * @throws IllegalArgumentException ����ֵ������
	 */
	public ArrayList<Row> getRowListByArg(int startRowIndex,int length,String title,String value) throws
													IndexOutOfBoundsException,IllegalArgumentException{
		int count = rowList.size();
		try {
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
			ArrayList<Row> rows = new ArrayList<>();
			for(int i=startRowIndex;i<=endIndex;i++){
				rows.add(rowList.get(i));
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
	 * @throws IndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
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
			int length = rowList.size();
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
	 * @throws IndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 */
	public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = rowList.size();
		try {
			int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
			Cell cell = null;
			Row row = null;
			for(int i=startRowIndex;i<=endIndex;i++){
				row = rowList.get(i);
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
	 * @param rowIndex sheet�����±�
	 * @param startColumnIndex ��ʼ���±�ֵ
	 * @param length ���賤��
	 * @throws IndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 * @throws IllegalArgumentException ָ����Ϊnull,��û������
	 */
	public void setTitles(int rowIndex,int startColumnIndex,int length) throws IndexOutOfBoundsException,
																IllegalArgumentException{
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			throw new IllegalArgumentException("����Ϊnull");
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
	 * <br>�������͸�ʽ(yyyy-MM-dd,yyyy-MM,MM-dd),��Ԫ��ʽΪDATE_NUMʱֻ����(yyyy-MM-dd)
	 * <br>cellΪnull,��û��ֵ,��ȡֵ�����򷵻�""
	 */
	public static String getCellValue(Cell cell){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			TypeEnum cellType = CellTypeUtil.getCellType(cell);
			
			switch (cellType) {
			case STRING:
				string =  cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				double dd = cell.getNumericCellValue();
				int in = (int) dd;
				double last = dd - in;
				if (last == 0) // doubleΪ����
					string = String.valueOf(in);
				else
					string = String.valueOf(dd);
				break;
			case DATE_NUM:
				double d = cell.getNumericCellValue();
				Date date = DateUtil.getJavaDate(d);
				string = new SimpleDateFormat("yyyy-MM-dd").format(date);
				break;
			case DATE_STR:
				string =  cell.getStringCellValue().trim();
				switch (CellTypeUtil.getDateEnum(string)) {
				case yyyy_MM_dd_chinese:
					string = string.replaceAll("[����]{1}", "-");
					string = string.replaceAll("[�պ�]?", "");
					break;
				case yyyy_MM_chinese:
					string = string.replaceAll("[����]{1}", "-");
					break;
				case MM_dd_chinese:
					string = string.replaceAll("��{1}", "-");
					string = string.replaceAll("[�պ�]?", "");
					break;
				case yyyy_MM_dd:
					string = string.replaceAll("([/-]|\\.){1}", "-");
					break;
				case yyyy_MM:
					string = string.replaceAll("[/-]{1}", "-");
					break;
				case MM_dd:
					string = string.replaceAll("[/-]{1}", "-");
					break;		
				default:
					break;
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
			//ȥ����βȫ�ǿհ׷�
			while (string.startsWith("��")) {
				string = string.substring(1, string.length()).trim();
			}
			while (string.endsWith("��")) {
				string = string.substring(0, string.length() - 1).trim();
			}
			/*if (string != null) {//ȥ�����С��س����Ʊ��  
		        Pattern p = Pattern.compile("\\s*|\t|\r|\n");  
		        Matcher m = p.matcher(string);  
		        string = m.replaceAll("");  
		    }  */
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
	 * ǿ�ƽ�Numeric���͵�ֵת��Ϊ����
	 * @param numeric numeric��ʽ�ĵ�Ԫ����ֵ
	 * @return yyyy-MM-dd��ʽ������
	 */
	public static String changeNumericToDate(double numeric){
		String string = "";
		Date date = DateUtil.getJavaDate(numeric);
		string = new SimpleDateFormat("yyyy-MM-dd").format(date);
		return string;
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
	 * @return ����Ҫ��ȡ�����һ���±�ֵ(����±�Խ��,�򷵻�����±�ֵ)
	 * <br>�������Խ���׳��쳣
	 * @throws IndexOutOfBoundsException ��������
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
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
