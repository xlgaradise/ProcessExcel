
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
 *读sheet表的工具,可获取行、单元格等
 */
public class SheetReadXMLUtil {
	
	protected InputSource source;
	protected XMLReader xmlReader;
	protected InputStream sheetInputStream;
	protected MyXSSFSheetHandler handler;
	
	/**
	 * 自选标题(Integer为标题的列下标,String为标题内容)
	 */
	protected HashMap<Integer, MyCell> title;
	
	/**
	 * （从ExcelReadXMLUtil获取实例）
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
	 * 解析指定范围内的数据
	 * @param rowsLength 读取的行数（base 1）
	 * @param colsLength 读取的列数（base 1）
	 * @throws ExcelIllegalArgumentException 参数不能小于1
	 * @throws ExcelFileOpenException 文件读取错误
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
     * 获取单元格
     * @param rowIndex 行下标(base 0)
     * @param colIndex 列下标(base 0)
     * @return 返回实例或者null
     * @throws ExcelIndexOutOfBoundsException 下标小于0或者大于读取值
     */
	public MyCell getCell(int rowIndex,int colIndex) throws ExcelIndexOutOfBoundsException{
		return handler.getCell(rowIndex, colIndex);
	}
	
	/**
     * 获取单元格值
     * @param rowIndex 行下标(base 0)
     * @param colIndex 列下标(base 0)
     * @return 有值返回其值 否则返回 ""
     * @throws ExcelIndexOutOfBoundsException 下标小于0或者大于读取值
     */
	public String getCellValue(int rowIndex,int colIndex) throws ExcelIndexOutOfBoundsException{
		MyCell cell = getCell(rowIndex, colIndex);
		return  cell != null ? cell.value : "";
	}
	
	/**
     * 判断指定下标行是否为null
     * @param rowIndex 行下标(base 0)
     * @return 
     * @throws ExcelIndexOutOfBoundsException 下标值小于0或者大于读取值
     */
	public boolean isRowIsNull(int rowIndex) throws ExcelIndexOutOfBoundsException{
		return handler.isRowIsNull(rowIndex);
	}
	
	/**
	 * 返回读取的所有行
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
     * 返回最后有值的行所在下标+1
     * @return
     */
	public int getRowsSize(){
		return handler.getRowsSize();
	}
	
	/**
	 * 设定sheet表的标题
	 * @param rowIndex 标题所在行下标 (base 0)
	 * @param startColumnIndex 起始列下标值 (base 0)
	 * @param length 标题长度 (base 1)
	 * @throws ExcelIndexOutOfBoundsException 下标值小于0或者大于读取值
	 * @throws ExcelIllegalArgumentException  指定行没有数据
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
	 * 通过标题值获取所在列下标
	 * @param titleName 指定标题值
	 * @return 标题列下标,不存在则返回-1
	 * @throws ExcelNoTitleException 未设置标题
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
	 *以XML形式解析Excel文件
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
        	if(name.equals("row")){//新的一行
        		oneRow = new HashMap<>();
        	}
            else if ("c".equals(name)) {//单元格开始
            	
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

                /*-----补缺空余行-----*/
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
            else if(name.equals("f")){//formula 包含在v下
            	value.delete(0, value.length());
            }
            else if(name.equals("v")){//value  包含在v下
            	value.delete(0, value.length());
            }
  
        }  
  
        public void endElement(String uri, String localName, String name)  
                throws SAXException {  
        	
        	if(lastRowIndex == rowLimit) return;
        	
            String valueStr = "";  
            if(name.equals("c")){//单元格结束
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
					// 判断是否是日期格式
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
								currentCell.formatIndex ==179) {// 自定义日期格式
							/**
							 * Excel中自定义日期类型 (yyyy/m/d) format:14 (yyyy年MM月dd日) format:31
							 * (yyyy年MM月) format:57 (MM-dd或MM月dd日) format:58 其他暂不确定
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
			else if (name.equals("row")) {// 行结束
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
         * 初始化参数
         * @param rowLimit 
         * @param colLimit
         * @throws ExcelIllegalArgumentException 参数小于1
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
         * 获取单元格
         * @param rowIndex 行下标(base 0)
         * @param colIndex 列下标(base 0)
         * @return 返回实例或者null
         * @throws ExcelIndexOutOfBoundsException 下标小于0或者大于读取值
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
         * 返回读取所有行
         * @return
         */
        public ArrayList<HashMap<Integer, MyCell>> getRowsList(){
        	return rows;
        }
        
        /**
         * 返回最后有值的行所在下标+1
         * @return
         */
        public int getRowsSize(){
        	return rows.size();
        }
        
        /**
         * 判断指定下标行是否为null
         * @param rowIndex 行下标(base 0)
         * @return 
         * @throws ExcelIndexOutOfBoundsException 下标值小于0或者大于读取值
         */
        public boolean isRowIsNull(int rowIndex) throws ExcelIndexOutOfBoundsException{
        	if(rowIndex > rowLimit || rowIndex < 0) 
        		throw new ExcelIndexOutOfBoundsException();
        	if(rowIndex > rows.size()-1) return true; 
        	HashMap<Integer, MyCell> row = rows.get(rowIndex);
        	return row == null ? true : false;
        }
  
        
        /**
         * 将列名称转换为列下标
         * @param name 例如 ‘BC’
         * @return 下标值（base 0）
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
