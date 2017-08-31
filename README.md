# Azure Excel Add-in
An Excel add-in for accessing Azure functions like usage aggregation reports (consumption) and rate card for standard Azure accounts, [Enterprise Agreement](https://www.microsoft.com/en-us/licensing/licensing-programs/enterprise.aspx) (EA) accounts, and [Cloud Solution Provider](https://partner.microsoft.com/en-US/cloud-solution-provider) (CSP) accounts.

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
* Enabling API Access:
> In order to be able to access the EA billing APIs programmatically, you must go to the EA portal, https://ea.azure.com, and generate an API key, as described on the following page:
https://docs.microsoft.com/en-us/azure/billing/billing-enterprise-api

### Cloud Solution Provider (CSP)
* Usage: https://msdn.microsoft.com/en-us/library/partnercenter/mt791774.aspx
* Ratecard: https://msdn.microsoft.com/en-us/library/partnercenter/mt774619.aspx
* Invoice: https://msdn.microsoft.com/en-us/library/partnercenter/mt712733.aspx
* Enabling API Access:
> In order to be able to access the CSP billing APIs programmatically, you must go to the Partner Center portal and enable API access, as described on the following page: https://msdn.microsoft.com/library/partnercenter/mt709136.aspx. Please note that bullet item #2 under "*Enable API access*" on that page incorrectly states: "*From the Dashboard menu, select Account Settings, then __API__.*" Instead, it should say: "*From the Dashboard menu, select Account settings, then __App Management__.*"

