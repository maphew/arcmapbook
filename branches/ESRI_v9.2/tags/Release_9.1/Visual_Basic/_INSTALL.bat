REM register component
regsvr32 /s ".\DSMapBookPrj.dll"
regsvr32 /s ".\DSMapBookUIPrj.dll"

REM register components in appropriate component categories
.\register_component_category.reg
