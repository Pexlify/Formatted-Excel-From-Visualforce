public with sharing class SimpleOpportunityExcelController {
	
    public List<Opportunity> AllOpenOpportunities {get; private set;} 

    public SimpleOpportunityExcelController() {
		AllOpenOpportunities = [SELECT Id, Type, LeadSource, Amount, ExpectedRevenue, CloseDate, StageName, CreatedDate 
                            FROM Opportunity 
                            WHERE StageName != 'Closed Won'
                            LIMIT 2];
	}
}