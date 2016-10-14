clear
clear matrix
capture log close
set more off

log using "log/MS.log", replace

use "data/KPI"

global figures "reports"

*TABS*

tab leveld, missing
tab functiontype, missing
tab attainment, missing
tab executive, missing
tab region, missing
tab date, missing

bysort region: tab leveld
bysort region: tab functiont

*SUM*

bysort leveld: sum centrality
bysort functiont: sum centrality

bysort leveld: sum internalcoll
bysort leveld: sum externalcoll
bysort functiont: sum externalcoll
bysort functiont: sum internalcoll

bysort leveld: sum meetinghoursbyisexternal_true
bysort leveld: sum meetinghoursbyisexternal_false
bysort functiont: sum meetinghoursbyisexternal_true
bysort functiont: sum meetinghoursbyisexternal_false

bysort leveld: sum externalnetworksize
bysort functiont: sum internalnetworksize

*CORRS*

corr meetinghours  meeting_hours_high_quality
corr meetinghours meetingsattended
corr meetingsattended attendedmeetingquality

corr  meetingsattended  redundant
corr  attendedmeetingquality  redundant
corr  meeting_hours_high_quality  redundant

corr  meetinghours  centrality
corr  mailhours  centrality
corr  overload  centrality
corr  overloadmailscount  centrality
corr  overloadmeetings  centrality
corr  receivedmails  centrality
corr  sentmails  centrality
corr  totalhours  centrality
    
corr  organizedmeetinghours  organizedmeetingquality

corr  attend*quality meetingbreakdownrecurring_nonrec
corr  attend*quality meetingbreakdownrecurring_rec

bysort functiontype: corr  attend*quality meetingbreakdownrecurring_nonrec
bysort functiontype: corr  attend*quality meetingbreakdownrecurring_rec

bysort leveld: corr  attend*quality meetingbreakdownrecurring_nonrec
bysort leveld: corr  attend*quality meetingbreakdownrecurring_rec

bysort leveld: corr meetinghours mailhours
bysort functiont: corr meetinghours mailhours
bysort region: corr meetinghours mailhours

corr meetinghours mailhours

*REGRESSIONS*

label var attendedmeetingquality "Attended Mtg Quality - 1 to 100 Scale"

regress attendedmeetingquality function1-function6, cluster(pid)
outreg2 using "$figures/regression_basic_$S_DATE", excel label replace
regress attendedmeetingquality level1-level6, cluster(pid)
outreg2 using "$figures/regression_basic_$S_DATE", excel label append

foreach var in totalhours meetinghours mailhours overload overloadmeeting overloadmails doublebooked lowengage {
regress `var' function1-function6, cluster(pid)
outreg2 using "$figures/regression_basic_$S_DATE", excel label append
regress `var'  level1-level6, cluster(pid)
outreg2 using "$figures/regression_basic_$S_DATE", excel label append
}

regress attendedmeetingquality meetinghours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label replace
regress attendedmeetingquality mailhours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality totalhours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality overloadmeeting region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality overloadmail region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality doublebooked region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality lowenga region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality meetingbreakdownduration* region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append
regress attendedmeetingquality meetingbreakdownatte* region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "$figures/regression_quality_$S_DATE", excel label append

regress meetinghours mailhours, cluster(pid)
outreg2 using "$figures/regression_$S_DATE", excel label replace

levelsof leveld, local(levels) 
foreach l of local levels {
regress meetinghours mailhours if leveld=="`l'", cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
*regress meetinghours mailhours region2-region6 if leveld=="`l'", cluster(pid)
*outreg2 using "$figures/regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
}

levelsof functiont, local(levels) 
foreach l of local levels {
regress meetinghours mailhours if functiont=="`l'", cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
*regress meetinghours mailhours region2-region6 if functiont=="`l'", cluster(pid)
*outreg2 using "$figures/regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
}

regress meetinghours mailhours level2-level7, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append
regress meetinghours mailhours level2-level7 region2-region6, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append
regress meetinghours mailhours function2-function7, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append
regress meetinghours mailhours function2-function7 region2-region6, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append
regress meetinghours mailhours level2-level7 function2-function7 region2-region6, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append

regress central meetinghours mailhours, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append
regress central meetinghours mailhours level2-level7 function2-function7 region2-region6, cluster(pid)
outreg2 using "$figures/regression_hours_$S_DATE", excel label append

*FIGURES*

*relabeling some variables so the text fits better on the figures*

label var attendedmeetingquality "A score 0 - 100 of avg quality of mtgs attended"
label var organizedmeetingquality "A score 0 - 100 of avg quality of mtgs organized"

label var meetingbreakdownattendees_at "Hours spent in mtgs w/ 21+"
label var meetingbreakdownattendees_from1 "Hours spent in mtgs w/ 11-20"
label var meetingbreakdownattendees_from3 "Hours spent in mtgs w/ 3-6"
label var meetingbreakdownattendees_from7 "Hours spent in mtgs w/ 7-10"
label var meetingbreakdownattendees_only1 "Hours spent in mtgs w/ 1"
label var meetingbreakdownattendees_only2 "Hours spent in mtgs w/ 2"
label var mailhours "Hours spent on email"

label var meetingbreakdownduration_atl "Hours in 8+ hour mtgs"
label var meetingbreakdownduration_from15 "Hours in 15-29 min mtgs"
label var meetingbreakdownduration_from1to "Hours in 1-2 hour mtgs"
label var meetingbreakdownduration_from2 "Hours in 2-8 hour mtgs"
label var meetingbreakdownduration_from30 "Hours in 30-59 min mtgs"
label var meetingbreakdownduration_lesstha "Hours in less 15 min mtgs"
label var meetingbreakdownrecurring_n "Hours in non-recurring mtgs."
label var meetingbreakdownrecurring_r "Hours in recurring mtgs."

label var centrality "How central person is to flow of info w/in the company"
label var utilization "Total for wk, hrs btw first/last sent email or mtg attended"

foreach var in centrality  {
graph bar (mean) `var', over(leveld, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(`var'.png) replace    
erase `var'.png
erase `var'.gph  
}

foreach var in meetinghours meetingsattended attendedmeetingquality  mailhours sentmails receivedmails timeblock_1hr timeblock_2hrs doublebooked utilization internalnetworksize externalnetworksize {
graph bar (mean) `var', over(leveld, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}

foreach var in centrality  meetinghours meetingsattended attendedmeetingquality  mailhours sentmails receivedmails timeblock_1hr timeblock_2hrs doublebooked utilization internalnetworksize externalnetworksize {
graph bar (mean) `var', over(functiont, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}

*relabeling some variables so the text fits better on the figures*

label var meetingbreakdownduration_lesstha  "<15min"
label var meetingbreakdownduration_from15   "15-30min"
label var meetingbreakdownduration_from30t  "30min - 1hr"
label var meetingbreakdownduration_from1to  "1-2 hr"
label var meetingbreakdownduration_from2to  "2-8 hr"
label var meetingbreakdownduration_atleast  "8hr+"

replace leveld="Director" if leveld=="1 Director"
replace leveld="Senior Executive" if leveld=="2 Senior Executive"
replace leveld="Executive" if leveld=="3 Executive"
replace leveld="Manager" if leveld=="4 Manager"
replace leveld="Senior IC" if leveld=="5 Senior IC"
replace leveld="Junior IC" if leveld=="6 Junior IC"
replace leveld="Support" if leveld=="7 Support"

*replacing the "total hours by duration" with "share of total hours by duration"*

replace meetingbreakdownduration_lesstha=meetingbreakdownduration_lesstha/meetinghours
replace meetingbreakdownduration_from15=meetingbreakdownduration_from15/meetinghours
replace meetingbreakdownduration_from30t=meetingbreakdownduration_from30t/meetinghours
replace meetingbreakdownduration_from1to=meetingbreakdownduration_from1to/meetinghours
replace meetingbreakdownduration_from2to=meetingbreakdownduration_from2to/meetinghours
replace meetingbreakdownduration_atleast=meetingbreakdownduration_atleast/meetinghours

levelsof leveld, local(levels) 
foreach l of local levels {
graph bar meetingbreakdownduration_lesstha  meetingbreakdownduration_from15 meetingbreakdownduration_from30t meetingbreakdownduration_from1to meetingbreakdownduration_from2to meetingbreakdownduration_atleast if leveld=="`l'", subtitle("Share of Total Meeting Hours by Duration, for `l'") ytitle("Share of Total Meeting Hours") legend(lab(1 "<15min") lab(2 "15-30min") lab(3 "30min - 1hr") lab(4 "1-2 hr") lab(5 "2-8 hr") lab(6 "8hr+") symxsize(5) row(1)) 
graph save "`l'", replace
graph export "`l'.png", replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g("`l'.png") append        
erase "`l'.png"
erase "`l'.gph"  
}

levelsof functiont, local(levels) 
foreach l of local levels {
graph bar meetingbreakdownduration_lesstha  meetingbreakdownduration_from15 meetingbreakdownduration_from30t meetingbreakdownduration_from1to meetingbreakdownduration_from2to meetingbreakdownduration_atleast if functiont=="`l'", subtitle("Share of Total Meeting Hours by Duration, for `l'") ytitle("Share of Total Meeting Hours") legend(lab(1 "<15min") lab(2 "15-30min") lab(3 "30min - 1hr") lab(4 "1-2 hr") lab(5 "2-8 hr") lab(6 "8hr+") symxsize(5) row(1)) 
graph save "`l'", replace
graph export "`l'.png", replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g("`l'.png") append        
erase "`l'.png"
erase "`l'.gph"  
}

*replacing the "total hours by attendees" with "share of total hours by attendees"*

replace meetingbreakdownattendees_only1=meetingbreakdownattendees_only1/meetinghours
replace meetingbreakdownattendees_only2  =meetingbreakdownattendees_only2  /meetinghours
replace meetingbreakdownattendees_from3t=meetingbreakdownattendees_from3t/meetinghours
replace meetingbreakdownattendees_from7t=meetingbreakdownattendees_from7t/meetinghours
replace meetingbreakdownattendees_from11=meetingbreakdownattendees_from11/meetinghours
replace meetingbreakdownattendees_atleas=meetingbreakdownattendees_atleas/meetinghours

levelsof leveld, local(levels) 
foreach l of local levels {
#delimit ;
graph bar meetingbreakdownattendees_only1  meetingbreakdownattendees_only2  meetingbreakdownattendees_from3t  meetingbreakdownattendees_from7t  meetingbreakdownattendees_from11  meetingbreakdownattendees_atleas if leveld=="`l'",
subtitle("Share of Total Meeting Hours by Meeting Size, for `l'") ytitle("Share of Total Meeting Hours") legend(lab(1 "1 other") lab(2 "2 others") lab(3 "3-6") lab(4 "7-10") lab(5 "11-20") lab(6 "21+") symxsize(5) row(1)); 
#delimit cr
graph save "`l'", replace
graph export "`l'.png", replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g("`l'.png") append        
erase "`l'.png"
erase "`l'.gph"  
}

levelsof functiont, local(levels) 
foreach l of local levels {
#delimit ;
graph bar meetingbreakdownattendees_only1  meetingbreakdownattendees_only2  meetingbreakdownattendees_from3t  meetingbreakdownattendees_from7t  meetingbreakdownattendees_from11  meetingbreakdownattendees_atleas if functiont=="`l'",
subtitle("Share of Total Meeting Hours by Meeting Size, for `l'") ytitle("Share of Total Meeting Hours") legend(lab(1 "1 other") lab(2 "2 others") lab(3 "3-6") lab(4 "7-10") lab(5 "11-20") lab(6 "21+") symxsize(5) row(1)); 
#delimit cr
graph save "`l'", replace
graph export "`l'.png", replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g("`l'.png") append        
erase "`l'.png"
erase "`l'.gph"  
}

*RECURRING MEETING ANALYSIS*

*replacing total hours by share of hours*
            
g share_recurring=meetingbreakdownrecurring_rec/meetinghours
g share_nonrecurring=meetingbreakdownrecurring_nonrec/meetinghours
g share_externalmtg=meetinghoursbyisexternal_true/meetinghours
g share_externalall=externalcoll/totalhours

*relabeling some variables so the text fits better on the figures*

label var share_recurring "Share of Meeting Hours Spent in Recurring Meetings"
label var share_nonrecurring "Share of Meeting Hours Spent in Non-Recurring Meetings"
label var share_externalmtg "Share of Meeting Hours Spent in External Meetings"
label var share_externalall "Share of Meeting/Email Hours External"

corr  attend*quality share_recurring
corr  attend*quality share_nonrecurring
bysort functiontype: corr  attend*quality share_recurring
bysort functiontype: corr  attend*quality share_nonrecurring
bysort leveld: corr  attend*quality share_recurring
bysort leveld: corr  attend*quality share_nonrecurring

foreach var in share_recurring share_externalmtg share_externalall {
graph bar (mean) `var', over(functiont, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}
		
foreach var in share_recurring share_externalmtg share_externalall {
graph bar (mean) `var', over(leveld, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}
		
