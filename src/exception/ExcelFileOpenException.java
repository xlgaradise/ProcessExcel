
package exception;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *文件打开错误
 */
public class ExcelFileOpenException extends ExcelBaseException {
	
	private Exception exception;

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ExcelFileOpenException(Exception exception){
		this.exception = exception;
	}
	
	public Exception getException(){
		return exception;
	}
}
