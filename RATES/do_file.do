clear all
cd "C:\Users\WB459082\Documents\RATES\"
import excel "appended_complete.xlsx", sheet("Sheet1") firstrow

replace POL = LOCATION if POL==""
replace POL = POLPOD if POL==""
drop if POL==""
drop LOCATION POLPOD A
replace POL = proper(POL)


rename C var_20GP 
rename G var_40GP
rename HQ var_40HQ
rename NOR var_40NOR
destring var_*, replace force

replace var_20GP=GP if var_20GP==.
replace var_40GP=D if var_40GP==.
replace var_40HQ=HC if var_40HQ==.
replace var_40NOR=Nor if var_40NOR==.

drop if var_20GP==.
drop GP D HC Nor


bysort File Sheet POL: gen count=_n
replace POL="Semarang2" if count==2
reshape long var_ , i(File Sheet POL) j(type) string
rename var_ Rate 
drop count
order File Sheet Description Startdate Enddate POL Rate
export excel using "appended_complete_clean.xlsx", replace  firstrow(variables)