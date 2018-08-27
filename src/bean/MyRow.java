
package bean;

import java.util.HashMap;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *�Զ��������ͣ������xml��ʽ��ȡ
 *
 */
public class MyRow {

	private HashMap<Integer, MyCell> row;
	public MyRow(HashMap<Integer, MyCell> row){
		this.row = row;
	}
	
	/**
	 * ����ָ���±굥Ԫ��
	 * @param cellIndex
	 * @return ����MyCellʵ������null
	 */
	public MyCell getCell(int cellIndex){
		return row.get(cellIndex);
	}
	
	/**
	 * ������Ч��Ԫ��ĸ���
	 * @return
	 */
	public int getValidCellSize(){
		return row.size();
	}
	
	public void print(int colSize){
		String rowStr = "{";
		String message = "";
		for(int i=0;i<colSize;i++){
			MyCell cell = this.row.get(i);
			if(cell == null) 
				message += "[] ";
			else
				message += "["+cell.value+"] ";
		}
		rowStr += message +"}";
		System.out.println(rowStr);
	}
}
