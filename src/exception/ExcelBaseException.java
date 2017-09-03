
package exception;

/**
*@auchor HPC
*@encoding GBK
*/

public class ExcelBaseException extends Exception {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ExcelBaseException(){
		
	}
	
//	protected String message = "";
//	
//	public ExcelBaseException(String message) {
//		this.message = message;
//	}
//
//	public String getMessage() {
//		return message;
//	}

	@Override
	public Throwable fillInStackTrace() {
		return this;
	}
}
