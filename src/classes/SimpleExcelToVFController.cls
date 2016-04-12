public with sharing class SimpleExcelToVFController {
    public final String XML_HEADER {get; private set;}
    public final String WORKBOOK_END {get; private set;}

	public SimpleExcelToVFController() {
        XML_HEADER = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
		WORKBOOK_END = '</Workbook>';
	}
}