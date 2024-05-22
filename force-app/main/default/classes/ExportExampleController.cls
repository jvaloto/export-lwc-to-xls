// class just to export any dummy records
public class ExportExampleController{

    @AuraEnabled
    public static Opportunity getOpportunity(){
        return
            [SELECT Id
                ,  Name
                , Account.Name
                , Amount
                , CloseDate
                , LeadSource
                , Owner.Name
                , Type
            FROM Opportunity
            LIMIT 1];
    }

    @AuraEnabled
    public static List<PricebookEntry> getProducts(){
        return
            [SELECT Id
                , Product2Id
                , Product2.Name
                , Product2.ProductCode
                , UnitPrice
            FROM PricebookEntry
            WHERE Pricebook2.IsStandard = true
            ORDER BY Product2.Name ASC
            LIMIT 10];
    }
               

}