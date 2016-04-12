public with sharing class FormattedExcelController {
    public final String XML_HEADER {get; private set;}
    public final String WORKBOOK_END {get; private set;}

    private Transient List<Opportunity> AllOpportunities;
    private Transient Map<String, Double> OpportunityStageAmountByStage;
    private Transient Map<String, Integer> OpportunityStageCount;
    private Double LargestStageAmount;
    public Map<String, StageDetail> OpportunityStagesWithFormatting {get; private set;}
    public Opportunity LargestOpportunity {get; private set;}
    

    public FormattedExcelController() {
        XML_HEADER = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
        WORKBOOK_END = '</Workbook>';
        initializeVariables();
        processAllOpportunitiesCalculatingTotalsAndLargestOpportunity();
        calculateLargestStageAmount();
        wrapAllStages();
    }

    private void initializeVariables() {
        LargestStageAmount = 0;
        OpportunityStageAmountByStage = new Map<String, Double>();
        OpportunityStageCount = new Map<String, Integer>();
        LargestOpportunity = new Opportunity(Name='', Amount= 0.00);
        AllOpportunities = [SELECT Id, Name, Amount, StageName FROM Opportunity LIMIT 10000];
        OpportunityStagesWithFormatting = new Map<String, StageDetail>();
    }

    private void processAllOpportunitiesCalculatingTotalsAndLargestOpportunity() {
        for (Opportunity each_opportunity : AllOpportunities) {
            addOpportunityAmountToOpportunityStageCalculations(each_opportunity);
            calculateHighestAmountOpportunity(each_opportunity);
            incrementOpportunityStageCountForOpportunity(each_opportunity);
        }
    } 

    private void addOpportunityAmountToOpportunityStageCalculations(Opportunity opportunity_to_process) {
        String opportunity_stage = opportunity_to_process.StageName;
        if (isOpportunityStageNameNotPresentInOpportunityStageAmountByStageProperty(opportunity_stage)) {
            OpportunityStageAmountByStage.put(opportunity_stage, 0.00);
            OpportunityStageCount.put(opportunity_stage, 0);
        }
        addOpportunityAmmountToOpportunityStage(opportunity_to_process);
    }

    private Boolean isOpportunityStageNameNotPresentInOpportunityStageAmountByStageProperty(String stage_name) {
        return ! OpportunityStageAmountByStage.containsKey(stage_name);
    }

    private void addOpportunityAmmountToOpportunityStage(Opportunity opportunity_to_process) {
        Double current_amount_for_stage_name = OpportunityStageAmountByStage.get(opportunity_to_process.StageName);
        Double ammended_amount_for_stage_name = current_amount_for_stage_name + opportunity_to_process.Amount;
        OpportunityStageAmountByStage.put(opportunity_to_process.StageName, ammended_amount_for_stage_name);
    }

    private void calculateHighestAmountOpportunity(Opportunity opportunity_to_check) {
        Decimal opportunity_amount = opportunity_to_check.Amount;
        if (isAmountGreaterThanLargestOpportunityAmount(opportunity_amount)) {
            LargestOpportunity = opportunity_to_check;
        }
    }

    private Boolean isAmountGreaterThanLargestOpportunityAmount(Decimal opportunity_amount) {
        return opportunity_amount > LargestOpportunity.Amount; 
    }

    private void incrementOpportunityStageCountForOpportunity(Opportunity opportunity_to_process) {
        Integer current_quantity_for_stage_name = OpportunityStageCount.get(opportunity_to_process.StageName);
        Integer incremented_quantity_for_stage_name = current_quantity_for_stage_name + 1;
        OpportunityStageCount.put(opportunity_to_process.StageName, incremented_quantity_for_stage_name);
    }

    private void calculateLargestStageAmount() {
        for (String each_stage_name : OpportunityStageAmountByStage.keySet()) {
            if (OpportunityStageAmountByStage.get(each_stage_name) > LargestStageAmount) {
                LargestStageAmount = OpportunityStageAmountByStage.get(each_stage_name);
            }
        }
    }

    private void wrapAllStages() {
        for (String each_stage_name : OpportunityStageAmountByStage.keySet()) {
            OpportunityStagesWithFormatting.put(each_stage_name, new StageDetail(each_stage_name, this));
        }
    }

    public Set<String> getAllOpportunityStageHeadings() {
        return OpportunityStagesWithFormatting.keySet();
    }

    public class StageDetail {
        private Transient FormattedExcelController ParentClass;
        private Transient Boolean IsFirst {get; private set;}
        public String Name {get; private set;}
        public Integer Total {get; private set;}
        public Integer Quantity {get; private set;}
        public String Formatting {get; private set;}
        public String Formula {get; private set;}

        
        public StageDetail(String stage_name, FormattedExcelController calling_class) {
            ParentClass = calling_class;
            Name = stage_name;
            calculateIfIsFirst();
            setQuantity();
            setIntegerTotalBasedOnDecimalAmount();
            calculateFormattingForStageCell();
            calculateFormula();
        }

        private void calculateIfIsFirst() {
            Integer size_of_OpportunityStagesWithFormatting = ParentClass.OpportunityStagesWithFormatting.size();
            if (size_of_OpportunityStagesWithFormatting == 0) {
                IsFirst = true;
            } else {
                IsFirst = false;
            }
        }

        private void setQuantity() {
            Integer stage_quantity = ParentClass.OpportunityStageCount.get(Name);
            Quantity = stage_quantity;
        }

        private void setIntegerTotalBasedOnDecimalAmount() {
            Decimal stage_total = ParentClass.OpportunityStageAmountByStage.get(Name);
            Total = Integer.valueOf(stage_total);
        }

        private void calculateFormattingForStageCell() {
            if (Total >= ParentClass.LargestStageAmount) {
                Formatting = 's82';
            } else {
                Formatting = 's81';
            }
        }

        private void calculateFormula() {
            if (IsFirst) {
                Formula = '=R[-2]C';
            } else {
                Formula = '=R[-2]C+RC[-1]';
            }
        }
    }
}