# Azure Excel Add-in
An Excel add-in for accessing Azure functions like usage aggregation reports (consumption) and rate card for standard Azure accounts, Enterprise Agreement (EA) accounts, and Cloud Service Provider (CSP) accounts.

The add-in can be installed from the following location (ClickOnce installer updated by a CI build):
http://billingtools.azurewebsites.net/excel/install/setup.exe

## Add-in Installation Prerequisites:
* Windows 10 (might work on Windows 8.1 or Windows 7 but not tested)
* Excel 2016 (might work with Excel 2013 but not tested)

## Development Prerequisites:
* Visual Studio 2015 or 2017 
  - Download VS Community: https://www.visualstudio.com/vs/community/
* Office Developer Tools for Visual Studio installed
  - https://www.visualstudio.com/vs/office-tools/

## Reference Information:
### Standard Azure Accounts
* Usage: https://msdn.microsoft.com/en-us/library/azure/mt219003.aspx
* Ratecard: https://msdn.microsoft.com/en-us/library/azure/mt219005.aspx
* Invoice Download: https://docs.microsoft.com/en-us/rest/api/billing/ 

### Enterprise Agreeement (EA)
* Usage: https://docs.microsoft.com/en-us/azure/billing/billing-enterprise-api-usage-detail
* Price Sheet: https://docs.microsoft.com/en-us/azure/billing/billing-enterprise-api-pricesheet
* Balance & Summary: https://docs.microsoft.com/en-us/azure/billing/billing-enterprise-api-balance-summary 
* Marketplace Store Charge: https://docs.microsoft.com/en-us/azure/billing/billing-enterprise-api-marketplace-storecharge

### Cloud Service Provider (CSP)
* Enabling API Access: https://msdn.microsoft.com/en-us/library/partnercenter/mt709136.aspx
* Usage: https://msdn.microsoft.com/en-us/library/partnercenter/mt791774.aspx
* Ratecard: https://msdn.microsoft.com/en-us/library/partnercenter/mt774619.aspx
* Invoice: https://msdn.microsoft.com/en-us/library/partnercenter/mt712733.aspx
