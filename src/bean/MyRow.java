
package bean;

import java.util.HashMap;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *自定义行类型，仅针对xml形式读取
 *
 */
public class MyRow {

	private HashMap<Integer, MyCell> row;
	public MyRow(HashMap<Integer, MyCell> row){
		this.row = row;
	}
	
	/**
	 * 返回指定下标单元格
	 * @param cellIndex
	 * @return 返回MyCell实例或者null
	 */
	public MyCell getCell(int cellIndex){
		return row.get(cellIndex);
	}
	
	/**
	 * 返回有效单元格的个数
	 * @return
	 */
	public int getValidCellSize(){
		return row.size();
	}
}
