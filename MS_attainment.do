clear
clear matrix
capture log close
set more off

log using "logs\attainment.log", replace

use "data\KPI"

global figures "reports"

keep if functiont=="Sales"

tab attain, gen(attaindum)
rename attaindum1 highdum
rename attaindum2 meddum
rename attaindum3 lowdum

/*
foreach var in meeting_hours_high_quality	meetingbreakdownrecurring_nonrec	meetingsattended	externalnetworkbreadth	meetinghoursbylevel_above	meetingbreakdownduration_from30t	meetingbreakdownduration_from2to	centrality	receivedmails	meetinghoursbylevel_below	meeting_hours_duration_one_hour_	doublebookedhours	meetingbreakdownattendees_from3t	meetingbreakdownrecurring_recurr	internalnetworksize	organizedmeetinghours	sentmailsbyisexternal_true	meetinghoursbyisexternal_false	meetingbreakdownattendees_only2	sentmailsbyisexternal_false	meetingbreakdownattendees_only1	totalhours	externalnetworksize	mailhours	meetingbreakdownattendees_atleas	meetinghours	redundanthours	internalnetworkbreadth	meetingbreakdownduration_lesstha	meetinghoursbyisexternal_true	attendedmeetingquality	meetingbreakdownattendees_from7t	overloadmeetings	meetingbreakdownduration_from1to	lowengagementhours	organizedmeetingquality	meetinghoursbylevel_same	timespentwithlevel_below	meetingbreakdownduration_atleast	meetingbreakdownduration_from15t	meetingbreakdownattendees_from11	overloadmailscount	sentmails	internalcollaborationhours	timeblock_1hr	externalcollaborationhours	overload	timeblock_2hrs	utilization	{
d `var'
*nothing over 30 corr*
*corr highdum `var'
*nothing over 20 corr*
corr lowdum `var'
}

collapse (sum) meeting_hours_high_quality	meetingbreakdownrecurring_nonrec	meetingsattended	externalnetworkbreadth	meetinghoursbylevel_above	meetingbreakdownduration_from30t	meetingbreakdownduration_from2to	centrality	receivedmails	meetinghoursbylevel_below	meeting_hours_duration_one_hour_	doublebookedhours	meetingbreakdownattendees_from3t	meetingbreakdownrecurring_recurr	internalnetworksize	organizedmeetinghours	sentmailsbyisexternal_true	meetinghoursbyisexternal_false	meetingbreakdownattendees_only2	sentmailsbyisexternal_false	meetingbreakdownattendees_only1	totalhours	externalnetworksize	mailhours	meetingbreakdownattendees_atleas	meetinghours	redundanthours	internalnetworkbreadth	meetingbreakdownduration_lesstha	meetinghoursbyisexternal_true		meetingbreakdownattendees_from7t	overloadmeetings	meetingbreakdownduration_from1to	lowengagementhours	meetinghoursbylevel_same	timespentwithlevel_below	meetingbreakdownduration_atleast	meetingbreakdownduration_from15t	meetingbreakdownattendees_from11	overloadmailscount	sentmails	internalcollaborationhours	timeblock_1hr	externalcollaborationhours	overload	timeblock_2hrs	utilization ///
	(mean) organizedmeetingq attendedmeetingq, by (pid *dum)

foreach var in meeting_hours_high_quality	meetingbreakdownrecurring_nonrec	meetingsattended	externalnetworkbreadth	meetinghoursbylevel_above	meetingbreakdownduration_from30t	meetingbreakdownduration_from2to	centrality	receivedmails	meetinghoursbylevel_below	meeting_hours_duration_one_hour_	doublebookedhours	meetingbreakdownattendees_from3t	meetingbreakdownrecurring_recurr	internalnetworksize	organizedmeetinghours	sentmailsbyisexternal_true	meetinghoursbyisexternal_false	meetingbreakdownattendees_only2	sentmailsbyisexternal_false	meetingbreakdownattendees_only1	totalhours	externalnetworksize	mailhours	meetingbreakdownattendees_atleas	meetinghours	redundanthours	internalnetworkbreadth	meetingbreakdownduration_lesstha	meetinghoursbyisexternal_true	attendedmeetingquality	meetingbreakdownattendees_from7t	overloadmeetings	meetingbreakdownduration_from1to	lowengagementhours	organizedmeetingquality	meetinghoursbylevel_same	timespentwithlevel_below	meetingbreakdownduration_atleast	meetingbreakdownduration_from15t	meetingbreakdownattendees_from11	overloadmailscount	sentmails	internalcollaborationhours	timeblock_1hr	externalcollaborationhours	overload	timeblock_2hrs	utilization	{
d `var'
*nothing over 30 corr*
*corr highdum `var'
*nothing over 20 corr*
corr lowdum `var'
}
*/

g time_trend=0
replace time_trend=1 if dt=="1/9/2011"
replace time_trend=2 if dt=="1/16/2011"
replace time_trend=3 if dt=="1/23/2011"
replace time_trend=4 if dt=="1/30/2011"
replace time_trend=5 if dt=="2/6/2011"
replace time_trend=6 if dt=="2/13/2011"
replace time_trend=7 if dt=="2/20/2011"
replace time_trend=8 if dt=="2/27/2011"
replace time_trend=9 if dt=="3/6/2011"
replace time_trend=10 if dt=="3/13/2011"
replace time_trend=11 if dt=="3/20/2011"
replace time_trend=12 if dt=="3/27/2011"
replace time_trend=13 if dt=="4/3/2011"
replace time_trend=14 if dt=="4/10/2011"
replace time_trend=15 if dt=="4/17/2011"
replace time_trend=16 if dt=="4/24/2011"
replace time_trend=17 if dt=="5/1/2011"
replace time_trend=18 if dt=="5/8/2011"
replace time_trend=19 if dt=="5/15/2011"
replace time_trend=20 if dt=="5/22/2011"
replace time_trend=21 if dt=="5/29/2011"
replace time_trend=22 if dt=="6/5/2011"
replace time_trend=23 if dt=="6/12/2011"
replace time_trend=24 if dt=="6/19/2011"


g a_number=0
replace a_number=1 if attain=="Low"
replace a_number=2 if attain=="Medium"
replace a_number=3 if attain=="High"

quietly tab organization, gen(orgdum)

bysort attainment: sum centrality
sum centrality
bysort attainment: sum totalhours
sum totalhours
bysort attainment: sum  externalcollaborationhours
sum  externalcollaborationhours


/*
/*
areg a_number level1-level6 time_trend, absorb(organization) cluster(pid)
outreg2 using "$figures/attainment_areg_$S_DATE", excel label replace 
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
areg a_number `var'  level1-level6 time_trend, absorb(organization)  cluster(pid)
outreg2 using "$figures/attainment_areg_$S_DATE", excel label append
}
*/

reg a_number level1-level6 time_trend orgdum*,  cluster(pid)
outreg2 using "$figures/attainment_areg_$S_DATE", excel label replace 
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
reg a_number `var'  level1-level6 time_trend orgdum*,   cluster(pid)
outreg2 using "$figures/attainment_areg_$S_DATE", excel label append
}

regress a_number level1-level6 region1-region5 time_trend,   cluster(pid)
outreg2 using "$figures/attainment_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
regress a_number `var'  level1-level6 region1-region5 time_trend,   cluster(pid)
outreg2 using "$figures/attainment_$S_DATE", excel label append
}


logit highdum level1-level6 time_trend orgdum*,   cluster(pid)
outreg2 using "$figures/attainment_logit_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
logit highdum `var'  level1-level6 time_trend orgdum*,   cluster(pid)
outreg2 using "$figures/attainment_logit_$S_DATE", excel label append
}

probit highdum level1-level6 time_trend orgdum*,   cluster(pid)
outreg2 using "$figures/attainment_probit_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
probit highdum `var'  level1-level6 time_trend orgdum*,   cluster(pid)
outreg2 using "$figures/attainment_probit_$S_DATE", excel label append
}
*/
collapse (sum) meeting_hours_high_quality	meetingbreakdownrecurring_nonrec	meetingsattended	externalnetworkbreadth	meetinghoursbylevel_above	meetingbreakdownduration_from30t	meetingbreakdownduration_from2to	receivedmails	meetinghoursbylevel_below	meeting_hours_duration_one_hour_	doublebookedhours	meetingbreakdownattendees_from3t	meetingbreakdownrecurring_recurr	internalnetworksize	organizedmeetinghours	sentmailsbyisexternal_true	meetinghoursbyisexternal_false	meetingbreakdownattendees_only2	sentmailsbyisexternal_false	meetingbreakdownattendees_only1	totalhours	externalnetworksize	mailhours	meetingbreakdownattendees_atleas	meetinghours	redundanthours	internalnetworkbreadth	meetingbreakdownduration_lesstha	meetinghoursbyisexternal_true		meetingbreakdownattendees_from7t	overloadmeetings	meetingbreakdownduration_from1to	lowengagementhours	meetinghoursbylevel_same	timespentwithlevel_below	meetingbreakdownduration_atleast	meetingbreakdownduration_from15t	meetingbreakdownattendees_from11	overloadmailscount	sentmails	internalcollaborationhours	timeblock_1hr	externalcollaborationhours	overload	timeblock_2hrs	utilization ///
	(mean) a_number organizedmeetingq attendedmeetingq centrality, by (pid *dum  *dum* level* region* )

bysort a_number: sum centrality
sum centrality
bysort a_number: sum totalhours
sum totalhours
bysort a_number: sum  externalcollaborationhours
sum  externalcollaborationhours

	/*
areg a_number level1-level6 time_trend, absorb(organization) 
outreg2 using "$figures/attainment_areg_collapse_$S_DATE", excel label replace 
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
areg a_number `var'  level1-level6 time_trend, absorb(organization)  
outreg2 using "$figures/attainment_areg_collapse_$S_DATE", excel label append
}
*/

reg a_number level1-level6  orgdum*  
outreg2 using "$figures/attainment_areg_collapse_$S_DATE", excel label replace 
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
reg a_number `var'  level1-level6  orgdum*
outreg2 using "$figures/attainment_areg_collapse_$S_DATE", excel label append
}

regress a_number level1-level6 region1-region5
outreg2 using "$figures/attainment_collapse_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
regress a_number `var'  level1-level6 region1-region5    
outreg2 using "$figures/attainment_collapse_$S_DATE", excel label append
}

logit highdum level1-level6  orgdum*
outreg2 using "$figures/attainment_logit_collapse_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
logit highdum `var'  level1-level6  orgdum*,   
outreg2 using "$figures/attainment_logit_collapse_$S_DATE", excel label append
}

probit highdum level1-level6  orgdum*
outreg2 using "$figures/attainment_probit_collapse_$S_DATE", excel label replace
foreach var in meetinghours attendedmeetingquality organizedmeetingquality meeting_hours_high_quality redundanthours lowengagementhours overloadmeeting totalhours overload utilization centrality receivedmails sentmails externalnetworksize internalnetworksize externalnetworkbreadth internalnetworkbreadth externalcollaborationhours internalcollaborationhours {
probit highdum `var'  level1-level6  orgdum*,   
outreg2 using "$figures/attainment_probit_collapse_$S_DATE", excel label append
}	
	
stop

stop

clear
*/
use "data\KPI"

*What function has attainment data?*
bysort functiont: tab attain

bysort leveld: tab functiont
bysort functiont: tab leveld

*note - level, attainment, region, function all stay the same through the entire time period*
keep pid date leveld functiont region attainment
duplicates drop
count
drop date
duplicates drop
sort pid
count if pid==pid[_n-1]


/*
graph bar, over(leveld, label(angle(45))) title("Function=All Functions")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc",  g(f1.png) replace    
erase f1.png
erase f1.gph  

graph bar, over(functiont, label(angle(45))) title("Function Breakdown")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append  
erase f1.png
erase f1.gph  


levelsof functiont, local(levels) 
foreach l of local levels { 
graph bar if functiont=="`l'", over(leveld, label(angle(45)))  title("Function=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}


levelsof region, local(levels) 
foreach l of local levels { 
graph bar if region=="`l'", over(leveld, label(angle(45)))  title("Region=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}
*/
keep if functiont=="Sales"
tab leveld
tab region
*tab organization

replace attainment="1 High" if attainment=="High"
replace attainment="2 Medium" if attainment=="Medium"
replace attainment="3 Low" if attainment=="Low"

graph bar, over(attainment, label(angle(45))) title("Sales=Attainment Breakdown")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc",  g(f1.png) append    
erase f1.png
erase f1.gph  
stop


graph bar, over(region, label(angle(45))) title("Sales=Region Breakdown")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc",  g(f1.png) append    
erase f1.png
erase f1.gph  

stop
levelsof leveld, local(levels) 
foreach l of local levels { 
graph bar if leveld=="`l'", over(attainment, label(angle(45)))  title("Level=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}

levelsof region, local(levels) 
foreach l of local levels { 
graph bar if region=="`l'", over(attainment, label(angle(45)))  title("Region=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}

levelsof attainment, local(levels) 
foreach l of local levels { 
graph bar if attainment=="`l'", over(leveld, label(angle(45)))  title("Attainment=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}

levelsof attainment, local(levels) 
foreach l of local levels { 
graph bar if attainment=="`l'", over(region, label(angle(45)))  title("Attainment=`l'")
graph save f1, replace
graph export f1.png, height(400) replace
png2rtf using "$figures/microsoft_$S_DATE.doc", g(f1.png) append
erase f1.png
erase f1.gph  
}
*/





stop

*No variation across time in the attainment*
/*
keep date pid attain
sort pid date
g counter=1
bysort pid: egen sum=sum(counter)

replace attain=lower(attain)

foreach var in high medium low {
g `var'share=0
replace `var'share=1 if attain=="`var'"
bysort pid: egen `var'sum=sum(`var'share)
replace `var'share=`var'sum/sum
}

keep pid *share
duplicates drop
*/



stop

foreach var in organization region functiontype leveldes {
bysort attain: tab `var'
}


stop
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
		
