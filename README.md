# PortalRecordsMoverConsole
## Port of the Portal Records Mover [XrmToolbox](http://XrmToolbox.com) plugin by Tanguy Touzard

Original project: [PortalRecordsMover](https://github.com/MscrmTools/MscrmTools.PortalRecordsMover/)  

This project is console application version of the  MscrmTools.PortalRecordsMover plugin included with the XrmToolbox.  

This is intended to be integrated into a DevOps solution for automated build and deployment.

Please refer to the sample portal_mover.cmd for a list of avaialable params. They are summarized below: 

## Command Line Options
* **/settings:**	Name of the settings file to be loaded.  See the ExportSettings.json for an example.  An empty /settings: argument will mean that only command line arguments will be used

The following command line parameter values will override any value specified in the settings JSON file.  Passing in only the argument will null out any value found in the settings file

* **/website:**	required - ID of the source parent website record for the export
* **/createdon:**	Select records created on or after this date
* **/modifiedon:**	Select records modified on or after this date

* **/datefilteroptions:**CreateAndModify	Specify whether to use Modified On or Created On or both when filtering records.
										Possible values: CreateOnly = 1, ModifyOnly, CreateAndModify. Either number or name will work

* **/batchcount:**	number of items for the retrieve multiple calls. Defaults to 10

* **/priordays:**	Select records using a range of days prior to the run date.  For example, 30 would set the created on and/or modified on values to Now.AddDays(-30);  
					If specifed, this overrides the crated on and modified on

* **/activeonly:**  Select only Active records

* **/targetenv:**	Full URL of the environment to which records are imported.  Ex. https://contoso.crm.dynamics.com
* **/importfile:**	Filename from which records will be read for import. Can use a mask for date/time: portal export {0:yyyy-MM-dd}.xml

* **/sourceenv:**	Full URL of the environment from which records are retreieved. Ex. https://microsoft.crm.dynamics.com
* **/exportfile:**	Filename to which records will be saved post export. Can use a mask for date/time: portal export {0:yyyy-MM-dd}.xml

* **/sourceuser:**  Username for the Source environment
* **/sourcepass:**  Password for the Source environment

* **/targetuser:**  Username for the Target environment
* **/targetpass:**  Password for the Target environment

Example: export only, using command line parameters and a config file
* PortalRecordsMover /priordays:5 /activeonly:true /sourceenv:https://my-dev.crm.dynamics.com
* PortalRecordsMover /priordays:5 /settings:exportonly.settings.json

Example: export and import, using command line parameters.  Export and import are the same, so this is a move from dev to test
* PortalRecordsMover /priordays:5 /activeonly:true /sourceenv:https://my-dev.crm.dynamics.com /targetenv:https://my-test.crm.dynamics.com /exportfile:portal-move{0:yyyy-MM-dd}.xml /importfile:portal-move{0:yyyy-MM-dd}.xml
