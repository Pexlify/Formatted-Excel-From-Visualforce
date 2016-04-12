public with sharing class SimpleExcelToVfCleanedController {
    public final String XML_HEADER {get; private set;}
    public final String WORKBOOK_END {get; private set;}

    private Map<String, String> ColumnHeadings {get; private set;}
    public List<Opportunity> AllOpenOpportunities {get; private set;} 

    public SimpleExcelToVfCleanedController() {
        XML_HEADER = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
        WORKBOOK_END = '</Workbook>';
        
        setupColumnHadings();
        queryAllOpportunities();
    }

    private void setupColumnHadings() {
        ColumnHeadings = new Map<String, String>();
        ColumnHeadings.put('Id', 'Opportunity ID');
        ColumnHeadings.put('Type', 'Type');
        ColumnHeadings.put('LeadSource', 'Lead Source');
        ColumnHeadings.put('Amount', 'Amount');
        ColumnHeadings.put('ExpectedRevenue', 'Expected Revenue');
        ColumnHeadings.put('CloseDate', 'Close Date');
        ColumnHeadings.put('StageName', 'Stage');
        ColumnHeadings.put('CreatedDate', 'Created Date');
    }

    private void queryAllOpportunities() {
            String full_query_string = getFullQueryString();
            System.debug(full_query_string);
            AllOpenOpportunities = Database.query(full_query_string);
    }
    
    private String getFullQueryString() {
        String query_string = 'SELECT '; 
        String fields_to_query = getQueryFields();
        String query_from_clause = ' FROM Opportunity WHERE StageName != \'Closed Won\' LIMIT 100';

        query_string += fields_to_query + query_from_clause;
        return query_string;
    }

    private String getQueryFields() {
        String fields_to_return = '';
        for (String each_key : ColumnHeadings.keyset()) {
            fields_to_return += each_key + ', ';
        }
        return fields_to_return.removeEnd(', ');
    }

    public List<String> getColumnHeadings() {
        return ColumnHeadings.values();
    }

    public List<String> getColumnApiKeys() {
        return new List<String>(ColumnHeadings.keyset());
    }
}