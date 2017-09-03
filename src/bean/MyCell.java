
package bean;

import excelUtil.CellTypeUtil.TypeEnum;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *自定义单元格类型，仅针对xml形式读取
 */
public class MyCell {

	public TypeEnum type;
	public String value;
	public String formulaValue;
	public int colIndex = 0;
	public int rowIndex = 0;
	
	public int formatIndex;
	public String formatString;
	
	
	public MyCell(int rowIndex,int colIndex){
		type = TypeEnum.NUMERIC;
		value = "";
		formulaValue = "";
		this.rowIndex = rowIndex;
		this.colIndex = colIndex;
		formatIndex = -1;
		formatString = null;
	}
	
}
