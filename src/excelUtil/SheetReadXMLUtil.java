
package excelUtil;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import bean.MyCell;
import bean.MyRow;
import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelFileOpenException;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNoTitleException;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *��sheet��Ĺ���,�ɻ�ȡ�С���Ԫ���
 */
public class SheetReadXMLUtil {
	
	protected InputSource source;
	protected XMLReader xmlReader;
	protected InputStream sheetInputStream;
	protected MyXSSFSheetHandler handler;
	
	/**
	 * ��ѡ����(IntegerΪ��������±�,StringΪ��������)
	 */
	protected HashMap<Integer, MyCell> title;
	
	/**
	 * ����ExcelReadXMLUtil��ȡʵ����
	 * @param stylesTable
	 * @param stringsTable
	 * @param sheetInputStream
	 * @throws ExcelFileOpenException
	 */
	public SheetReadXMLUtil(StylesTable stylesTable,  
            ReadOnlySharedStringsTable stringsTable, InputStream sheetInputStream) throws ExcelFileOpenException{
		this.sheetInputStream = sheetInputStream;
		source = new InputSource(sheetInputStream);
		handler = new MyXSSFSheetHandler(stylesTable, stringsTable);
		try {
	        xmlReader =  SAXParserFactory.newInstance().newSAXParser().getXMLReader();    
		} catch (Exception e) {
			throw new ExcelFileOpenException(e);
		}
		xmlReader.setContentHandler(handler);
		title = new HashMap<>();
	}
	
	/**
	 * ����ָ����Χ�ڵ�����
	 * @param rowsLength ��ȡ��������base 1��
	 * @param colsLength ��ȡ��������base 1��
	 * @throws ExcelIllegalArgumentException ��������С��1
	 * @throws ExcelFileOpenException �ļ���ȡ����
	 */
	public void readRows(int rowsLength,int colsLength) throws ExcelIllegalArgumentException,ExcelFileOpenException{
		try {
			handler.init(rowsLength, colsLength);
			xmlReader.parse(source);
		} catch (IOException | SAXException e) {
			throw new ExcelFileOpenException(e);
		}
	}
	
	/**
     * ��ȡ��Ԫ��
     * @param rowIndex ���±�(base 0)
     * @param colIndex ���±�(base 0)
     * @return ����ʵ������null
     * @throws ExcelIndexOutOfBoundsException �±�С��0���ߴ��ڶ�ȡֵ
     */
	public MyCell getCell(int rowIndex,int colIndex) throws ExcelIndexOutOfBoundsException{
		return handler.getCell(rowIndex, colIndex);
	}
	
	/**
     * ��ȡ��Ԫ��ֵ
     * @param rowIndex ���±�(base 0)
     * @param colIndex ���±�(base 0)
     * @return ��ֵ������ֵ ���򷵻� ""
     * @throws ExcelIndexOutOfBoundsException �±�С��0���ߴ��ڶ�ȡֵ
     */
	public String getCellValue(int rowIndex,int colIndex) throws ExcelIndexOutOfBoundsException{
		MyCell cell = getCell(rowIndex, colIndex);
		return  cell != null ? cell.value : "";
	}
	
	/**
     * �ж�ָ���±����Ƿ�Ϊnull
     * @param rowIndex ���±�(base 0)
     * @return 
     * @throws ExcelIndexOutOfBoundsException �±�ֵС��0���ߴ��ڶ�ȡֵ
     */
	public boolean isRowIsNull(int rowIndex) throws ExcelIndexOutOfBoundsException{
		return handler.isRowIsNull(rowIndex);
	}
	
	/**
	 * ���ض�ȡ��������
	 * @return
	 */
	public ArrayList<MyRow> getRowsList(){
		ArrayList<MyRow> list = new ArrayList<>();
		MyRow row = null;
		for(HashMap<Integer, MyCell> obj:handler.getRowsList()){
			row = new MyRow(obj);
			list.add(row);
		}
		return list;
	}
	
	/**
     * ���������ֵ���������±�+1
     * @return
     */
	public int getRowsSize(){
		return handler.getRowsSize();
	}
	
	/**
	 * �趨sheet��ı���
	 * @param rowIndex �����������±� (base 0)
	 * @param startColumnIndex ��ʼ���±�ֵ (base 0)
	 * @param length ���ⳤ�� (base 1)
	 * @throws ExcelIndexOutOfBoundsException �±�ֵС��0���ߴ��ڶ�ȡֵ
	 * @throws ExcelIllegalArgumentException  ָ����û������
	 */
	public void setTitle(int rowIndex,int startColumnIndex,int length) throws ExcelIndexOutOfBoundsException,
									ExcelIllegalArgumentException{
		title.clear();
		if(isRowIsNull(rowIndex)){
			throw new ExcelIllegalArgumentException();
		}else{
			MyCell myCell = null;
			for(int i=startColumnIndex,len=startColumnIndex+length;i<len;i++){
				myCell = getCell(rowIndex, i);
				if(myCell != null){
					title.put(i, myCell);
				}else{
					myCell = new MyCell(rowIndex, i);
					myCell.type = TypeEnum.STRING;
					title.put(i, myCell);
				}
			}
		}
	}
	
	public MyRow getTitle(){
		return new MyRow(title);
	}
	
	/**
	 * ͨ������ֵ��ȡ�������±�
	 * @param titleName ָ������ֵ
	 * @return �������±�,�������򷵻�-1
	 * @throws ExcelNoTitleException δ���ñ���
	 */
	public int getTitleColIndexByValue(String titleName) throws ExcelNoTitleException{
		if(title.isEmpty()){
			throw new ExcelNoTitleException();
		}
		for(Map.Entry<Integer, MyCell> entry : title.entrySet()){
			if(entry.getValue().value.equals(titleName)) 
				return entry.getKey();
		}
		return -1;
	}

	public void close() throws IOException{
		this.sheetInputStream.close();
		xmlReader = null;
		source = null;
		handler = null;
	}
	
	
	/**
	 *��XML��ʽ����Excel�ļ�
	 */
	private class MyXSSFSheetHandler extends DefaultHandler {  
		  
        private StylesTable stylesTable;  
        private ReadOnlySharedStringsTable sharedStringsTable;  
        private final DataFormatter formatter;  

        private ArrayList<HashMap<Integer,MyCell>> rows;
        private HashMap<Integer,MyCell> oneRow;
        private MyCell currentCell;
        private StringBuilder value;
        
        private int lastRowIndex;
        private String lastElement;
        
        private int rowLimit;
        private int colLimit;
        
        public MyXSSFSheetHandler(StylesTable styles,ReadOnlySharedStringsTable strings) {  
            this.stylesTable = styles;  
            this.sharedStringsTable = strings;  
            this.formatter = new DataFormatter();
            this.value = new StringBuilder();  
            this.rows = new ArrayList<>();
            this.oneRow = new HashMap<>();
            
            lastRowIndex = -1;
            rowLimit = 2999;//3000
            colLimit = 99;//100
        }  
  
        public void startElement(String uri, String localName, String name,  
                Attributes attributes) throws SAXException {  
        	
        	if(lastRowIndex == rowLimit) return;
        	
        	lastElement = name;
        	/*name: c-cell v-value f-formula  row-row */
        	if(name.equals("row")){//�µ�һ��
        		oneRow = new HashMap<>();
        	}
            else if ("c".equals(name)) {//��Ԫ��ʼ
            	
                // Get the cell reference like 'AB4'
                String r = attributes.getValue("r");  
                int firstDigit = -1;  
                for (int c = 0; c < r.length(); ++c) {  
                    if (Character.isDigit(r.charAt(c))) {  
                        firstDigit = c;  
                        break;  
                    }  
                }  
                int colIndex = colNameToColIndex(r.substring(0, firstDigit));
                int rowIndex = Integer.parseInt(r.substring(firstDigit,r.length()))-1;

                /*-----��ȱ������-----*/
                while((lastRowIndex+1) < rowIndex){
                	rows.add(null);
                	lastRowIndex++;
                }
                
                currentCell = new MyCell(rowIndex, colIndex);
                if(colIndex > colLimit) return;
                  
                String cellTypeStr = attributes.getValue("t");  
                String cellStyleStr = attributes.getValue("s");  
                
                if("s".equals(cellTypeStr)){//string
                	currentCell.type = TypeEnum.STRING;
                }
                else if("b".equals(cellTypeStr)){//boolean
                	currentCell.type = TypeEnum.BOOLEAN;
                }
                else if("e".equals(cellTypeStr)){//error
                	currentCell.type = TypeEnum.ERROR;
                }
                else if (cellStyleStr != null) { 
                    int styleIndex = Integer.parseInt(cellStyleStr);  
                    XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);  
                    int formatIndex = style.getDataFormat();  
                    String formatString = style.getDataFormatString();  
                    if (formatString == null)  
                        formatString = BuiltinFormats  
                                .getBuiltinFormat(formatIndex);  
                    currentCell.formatIndex = formatIndex;
                    currentCell.formatString = formatString;
                }                   
            }  
            else if(name.equals("f")){//formula ������v��
            	value.delete(0, value.length());
            }
            else if(name.equals("v")){//value  ������v��
            	value.delete(0, value.length());
            }
  
        }  
  
        public void endElement(String uri, String localName, String name)  
                throws SAXException {  
        	
        	if(lastRowIndex == rowLimit) return;
        	
            String valueStr = "";  
            if(name.equals("c")){//��Ԫ�����
            	if(currentCell == null || currentCell.colIndex > colLimit){
            		return;
            	}
            	if(currentCell.type == TypeEnum.NUMERIC && currentCell.value.equals("")){
            		currentCell.type = TypeEnum.STRING;
            	}
            	oneRow.put(currentCell.colIndex, currentCell);
            	currentCell = null;
            }
            else if (name.equals("v")) {  
            	if(currentCell == null || currentCell.colIndex > colLimit){
            		return;
            	}
				switch (currentCell.type) {
				case BOOLEAN:
					char first = value.charAt(0);
					valueStr = first == '0' ? "FALSE" : "TRUE";
					break;
				case ERROR:
					valueStr = value.toString();
					break;
				case STRING:
					String strIndex = value.toString();
					try {
						int idx = Integer.parseInt(strIndex);
						XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
						valueStr = rtss.toString();
					} catch (NumberFormatException ex) {
						valueStr = "";
					}

					try {
						valueStr = CellTypeUtil.getFormatDate(valueStr);
						currentCell.type = TypeEnum.DATE_STR;
					} catch (ExcelIllegalArgumentException e) {
					}
					break;
				case FORMULA:
					try {
						valueStr = formatter.formatRawCellContents(Double.parseDouble(value.toString()),
								currentCell.formatIndex, currentCell.formatString);
					} catch (NumberFormatException e) {
						valueStr = value.toString();
					}
					break;
				case NUMERIC:
					String num = value.toString();
					// �ж��Ƿ������ڸ�ʽ
					if (HSSFDateUtil.isADateFormat(currentCell.formatIndex, currentCell.formatString)) {
						try {
							Double d = Double.parseDouble(num);
							valueStr = CellTypeUtil.getFormatDate(d);
							currentCell.type = TypeEnum.DATE_NUM;
						} catch (NumberFormatException e) {
							valueStr = num;
						} catch (NullPointerException e) {
							valueStr = num;
						}
					}else{
						if (currentCell.formatIndex == 31 || currentCell.formatIndex == 57 || 
								currentCell.formatIndex == 58 || currentCell.formatIndex == 14 ||
								currentCell.formatIndex ==179) {// �Զ������ڸ�ʽ
							/**
							 * Excel���Զ����������� (yyyy/m/d) format:14 (yyyy��MM��dd��) format:31
							 * (yyyy��MM��) format:57 (MM-dd��MM��dd��) format:58 �����ݲ�ȷ��
							 */
							try {
								Double d = Double.parseDouble(num);
								valueStr = CellTypeUtil.getFormatDate(d);
								currentCell.type = TypeEnum.DATE_NUM;
							} catch (NumberFormatException e) {
								valueStr = num;
							} catch (NullPointerException e) {
								valueStr = num;
							}
						}else{
							valueStr = num;
						}
					}
					break;
				default:
					valueStr = "";
					break;
				}
                currentCell.value = valueStr.trim();
            } 
            else if (name.equals("f")) {
            	if(currentCell == null || currentCell.colIndex > colLimit){
            		return;
            	}
            	currentCell.formulaValue = value.toString();
        		if(currentCell.type == TypeEnum.NUMERIC){
        			currentCell.type = TypeEnum.FORMULA;
        		}
			}
			else if (name.equals("row")) {// �н���
				lastRowIndex++;
				rows.add(oneRow);
			}
        }  
 
        public void characters(char[] ch, int start, int length)  
                throws SAXException {  
        	if(lastElement.equals("f") || lastElement.equals("v")){
        		value.append(ch, start, length);
        	}
        }  
        
        /**
         * ��ʼ������
         * @param rowLimit 
         * @param colLimit
         * @throws ExcelIllegalArgumentException ����С��1
         */
        public void init(int rowLimit,int colLimit) throws ExcelIllegalArgumentException{
        	oneRow.clear();
        	rows.clear();
        	value.delete(0, value.length());
        	currentCell = null;
        	
            lastRowIndex = -1;
            lastElement = "";
            if(rowLimit < 1 || colLimit < 1)
            	throw new ExcelIllegalArgumentException();
            this.rowLimit = rowLimit-1;
            this.colLimit = colLimit-1;
        }
        
        /**
         * ��ȡ��Ԫ��
         * @param rowIndex ���±�(base 0)
         * @param colIndex ���±�(base 0)
         * @return ����ʵ������null
         * @throws ExcelIndexOutOfBoundsException �±�С��0���ߴ��ڶ�ȡֵ
         */
        public MyCell getCell(int rowIndex,int colIndex) throws ExcelIndexOutOfBoundsException{
        	if(rowIndex > rowLimit || colIndex > colLimit || rowIndex < 0 || colIndex < 0) 
        		throw new ExcelIndexOutOfBoundsException();
        	if(rowIndex > rows.size()-1) return null;
        	HashMap<Integer, MyCell> row = rows.get(rowIndex);
        	if(row != null){
        		MyCell cell = row.get(colIndex);
        		if(cell != null){
        			return cell;
        		}else{
        			return null;
        		}
        	}else{
        		return null;
        	}
        }
        
        /**
         * ���ض�ȡ������
         * @return
         */
        public ArrayList<HashMap<Integer, MyCell>> getRowsList(){
        	return rows;
        }
        
        /**
         * ���������ֵ���������±�+1
         * @return
         */
        public int getRowsSize(){
        	return rows.size();
        }
        
        /**
         * �ж�ָ���±����Ƿ�Ϊnull
         * @param rowIndex ���±�(base 0)
         * @return 
         * @throws ExcelIndexOutOfBoundsException �±�ֵС��0���ߴ��ڶ�ȡֵ
         */
        public boolean isRowIsNull(int rowIndex) throws ExcelIndexOutOfBoundsException{
        	if(rowIndex > rowLimit || rowIndex < 0) 
        		throw new ExcelIndexOutOfBoundsException();
        	if(rowIndex > rows.size()-1) return true; 
        	HashMap<Integer, MyCell> row = rows.get(rowIndex);
        	return row == null ? true : false;
        }
  
        
        /**
         * ��������ת��Ϊ���±�
         * @param name ���� ��BC��
         * @return �±�ֵ��base 0��
         */
        private int colNameToColIndex(String name) {  
            int column = -1;  
            for (int i = 0; i < name.length(); ++i) {  
                char c = name.charAt(i);  
                column = (column + 1) * 26 + c - 'A';  
            }  
            return column;  
        }  
  
    }  
}
