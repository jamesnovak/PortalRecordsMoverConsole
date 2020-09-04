:: command line options.  See the PortalRecordsMover.settings.xml for more details 
:: these values will override any value specified in the settings XML file 

:: /settings:	Name of the settings file to be loaded.
:: /website:	required - ID of the source parent website record for the export
:: /createdon:	Select records created on or after this date
:: /modifiedon:	Select records modified on or after this date

:: /datefilteroptions:CreateAndModify	Specify whether to use Modified On or Created On or both when filtering records.  
::										Possible values: CreateOnly = 1, ModifyOnly, CreateAndModify.  Either number or name will work

:: /batchcount:10	number of items for the retrieve multiple calls. Defaults to 10

:: /priordays:10	Select records using a range of days prior to the run date.  For example, 30 would set the created on and/or modified on values to Now.AddDays(-30);  
::					If specifed, this overrides the crated on and modified on

:: /activeonly:  Select only Active records

:: /targetenv:	Full URL of the environment to which records are imported.  Ex. https://contoso.crm.dynamics.com
:: /importfile:	Filename from which records will be read for import. Can use a mask for date/time: portal export {0:yyyy-MM-dd}.xml

:: /sourceenv:	Full URL of the environment from which records are retreieved. Ex. Ex. https://microsoft.crm.dynamics.com
:: /exportfile:	Filename to which records will be saved post export. Can use a mask for date/time: portal export {0:yyyy-MM-dd}.xml

:: /sourceuser:**  Username for the Source environment
:: /sourcepass:**  Password for the Source environment

:: /targetuser:**  Username for the Target environment
:: /targetpass:**  Password for the Target environment

:: /targetappid:**  App ID for the target environment
:: /targetappsecret:**  Client secret for the target environment

:: If the /exportfile: or /importfile: are not specifed, the related action will not occur

:: Passing in only the argument will null out any value found in the settings file.
:: Passing in an empty settings argument will mean that only command line arguments will be used

:: example: export only, even if target and import values are present in the configuration
:: PortalRecordsMover /priordays:5 /activeonly:true /sourceenv:https://contoso-dev.crm.dynamics.com /exportfile:export_{0:yyyy-MM-dd}.xml /targetenv: /importfile:  

:: example: configuration file only, but override the number of prior days
:: PortalRecordsMover /priordays:5 /settings:PortalRecordsMover.exportonly.settings.xml 