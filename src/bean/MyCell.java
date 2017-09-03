
package bean;

import excelUtil.CellTypeUtil.TypeEnum;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *�Զ��嵥Ԫ�����ͣ������xml��ʽ��ȡ
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
