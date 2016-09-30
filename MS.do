clear
clear matrix
capture log close
set more off

log using "microsoft_$S_DATE.log", replace

/*********************************************************************

9/29/16

To do:
1. Seperate the "cleaning" portion into its own file
2. Divide up program into small chunks (check Shapiro)
3. Write a shell script to run all the do files in order

*********************************************************************/

/*
*CLEANING THE DATA AND PREPARING FOR ANALYSIS*

*insheeting .csv. Renamed "Demo KPI Dump - All Dates_v2.csv" to "KPI.csv", no other changes to file were made*
insheet using "KPI.csv", names case

*dropping rows without variables*
drop if dt==""

*getting the day, month, year numbers out of the string version of the date and destringing, and creating a true date variable*
split dt, p("/")

forvalues x=1/3 {
destring dt`x', replace
}

rename dt1 month
rename dt2 day
rename dt3 year

g date=mdy(month, day, year)
format date %td

/*********************************************************************

labeling variables using definition from "KPI Definitions.xls"
starred out labels are for variables that don't exist in the dataset,
but exist in "KPI Definitions.xls."

*********************************************************************/

label var doublebookedhours "Hours Scheduled for Multiple Simultaneous Meetings"
*label var leadershipexposurecount "# of times members interacted with someone at least one level above, i.e. part of the same meeting"
*label var leadershipexposurehours "Hours in meetings with people who are at least one level above"
label var lowengagementhours "hours spent in meetings where they were not fully engaged, due to double-booking, redundant, sending 2+ emails" 
*label var lowengagementpercent "The ratio of Low Engagement Hours to the total of Time in Meetings."
*label var managementredundanthours "Hours that 3 unique layers within the same function join the same meeting."
*label var managerhourstogether "Hours with direct manager"
label var meetingbreakdownduration_atl "Hours in 8+ hour meetings"
label var meetingbreakdownduration_from15 "Hours in 15-29 min meetings"
label var meetingbreakdownduration_from1to "Hours in 1-2 hour meetings"
label var meetingbreakdownduration_from2 "Hours in 2-8 hour meetings"
label var meetingbreakdownduration_from30 "Hours in 30-59 min meetings"
label var meetingbreakdownduration_lesstha "Hours in less 15 min meetings"
label var meetingbreakdownrecurring_n "Hours in non-recurring meetings."
label var meetingbreakdownrecurring_r "Hours in recurring meetings."
label var meetinghours "Hours in meetings"
label var meetinghoursbyisexternal_false "Hours in meetings with no external attendees"
label var meetinghoursbyisexternal_true "Hours in meetings with external attendees"
*label var meetinghoursbyiscustomer_false "Hours in meetings with no customers"
*label var meetinghoursbyiscustomer_true "Hours in meetings with customers"
label var meetinghoursbylevel_above "Hours in meetings with person 1+ level above"
label var meetinghoursbylevel_below "Hours in meetings with person 1+ level below"
label var meetinghoursbylevel_same "Hours in meetings with person same level"
label var meetingsattended "# of meetings attended"
*label var meetingtimespentinactivity "Hours spent with a meeting activity"
*label var nummeetingpeopleinactivity "# of people with a meeting activity"
*label var nummeetingsinactivity "# of meeting with an activity"
*label var nummeetingswithattribute "# of meetings with attributes"
*label var oli_avgattendees "=oli_attendees / oli_num"
*label var oli_avgduration "=oli_duration / oli_num"
*label var oli_avgrecipients "=oli_recipients / oli_mails"
*label var oli_attendees "sum of internal attendees in the mtgs the person organized"
*label var oli_coattendees "sum of ALL attendees in the mtgs the person organized"
*label var oli_duration "total hours of meetings organized w at least 1 internal attendee, unadjusted double"
*label var oli_meetinghours "total hours of meetings organized w at least 1 internal attendee, adjusted"
*label var oli_num "# of meetings organized w at least 1 internal attendee"
*label var oneononecoachinghours "hours where the only other attendee is their direct report"
*label var oneononemanagercount "# meetings when only other attendee is manager"
*label var oneononemanagerhours "hours where the only other attendee is their direct manager"
label var organizedmeetinghours "Hours of meetings person organized, agnostic of double booking"
label var overloadmeetings "Hours spent in meetings that took place outside of 8am-5pm M-F"
*label var percentredundanthours "=managementRedundantHours / meetingHours"
label var redundanthours "hours the person was 'redundant' where redundancy is defined as the lowest layers of the meeting not including the top two layers within the same function as defined by hasGroupings"
*label var skiplevelexposurecount "# meetings the person has attended in the time period with any hasGroupings internal person who is at least 2 levels above that person"
*label var skiplevelexposurehours "Hours the person in meetings in the time period with any hasGroupings internal person who is at least 2 levels above that person"
*label var staffmeetingcount "# of meetings with only the person and 2 or more direct reports and no other internal attendees"
*label var staffmeetinghours "hours in  meetings with only the person and 2 or more direct reports and no other internal attendees"
*label var toomanymailshours "Hours the person spent in meetings during which the person sent at least 2 emails."
label var mailhours "Hours spent on email, estimated as 5 min/sent mail and 2.5 min/received mail"
*label var mailtimespentinactivity "Hours spent with an email activity"
*label var numattachments "Average number of attachments sent by members of group"
*label var nummailpeopleinactivity "# people with mail activity"
*label var nummailswithactivity "# mails with an activity"
*label var nummailswithattribute "# mails with an activity"
*label var oli_mailhours "Hours internal recipients took to read the mail the person send"
*label var oli_mails "# mails sent with at least 1 internal recipient"
*label var oli_recipients "# internal recipients that received emails the person sent"
label var overloadmailscount "# emails sent outside of work time"
*label var overloadmailspercent "share of emails sent outside work time"
label var receivedmails "# mails received per week"
label var sentmails "# mails sent per week"
label var sentmailsbyisexternal_false "# mails sent, no external recipients"
label var sentmailsbyisexternal_true "# mails sent, 1+ external recipients"
*label var sentmailsbyiscustomer_false "# mails sent, no customers"
*label var sentmailsbyiscustomer_true "# mails sent, 1+ customers"
label var externalcollaborationhours "Hours spent in meetings and emails with external people"
*label var externalcollaborationhourspct "=externalCollaborationHours / totalHours"
*label var insularitybyfunction "hours in meetings + mail that only contain people within the pid's Function grouping"
*label var insularitybylevel "hours in meetings + mail that only contain people within the pid's Level grouping"
*label var insularitybyorganization "hours in meetings + mail that only contain people within the pid's Organization grouping"
*label var insularitybyregion "hours in meetings + mail that only contain people within the pid's region grouping"
label var internalcollaborationhours "Hours the person spends in meetings and emails where there is at least one other internal person"
*label var intrateamtotalcount "# meetings and mails where the only attendees are the person's direct reports and themselves."
*label var intrateamtotalhours "Hours meetings and mails where the only attendees are the person's direct reports and themselves."
*label var managercentralizationcount "Within each group, the percentage of a team's emails and meetings where that team's manager is involved. "
*label var managercentralizationhours "Within each group, the total hours of a team's emails and meetings where that team's manager is involved."
*label var managerloadabsorption "Average ratio of a manager's meeting and mail hours to their team's meeting and mail hours."
*label var oli_hours "Hours consumed from the rest of the organization based on meetings you organized and emails you sent"
*label var oli_hoursreceived "Hours that other internal people consumed from you, based on the meetings they organized and the emails they sent"
label var overload "Hours spent on meetings and sent mail out of 8-5 MF"
*label var teamcollaborationhours "Hours of all collaboration time for a manager and each of their direct reports."
label var totalhours "Hours of meetings and email"
label var utilization "Utilization is calculated by looking at the spread of hours between a person's first sent email or meeting attended in a day and that person's last sent email or meeting attended that day."
*label var weekly_customer_domain_i "# of customers individual/group has 1+ interaction (1+ meeting or email<5 people)"
*label var weekly_customer_interactions "# of customer contacts individual/group has 1+ interaction  (1+ meeting or email<5 people)"
*label var weekly_external_domain_inter "# of external domains individual/group has 1+ interaction  (1+ meeting or email<5 people)"
*label var weekly_external_interactions "# interactions individual/group have with people outside the company  (1+ meeting or email<5 people)"
*label var weekly_internal_interactions "# interactions individual/group have with people within the company  (1+ meeting or email<5 people)"
*label var weekly_internal_network_velocity "# new people interacted with by individual/group  (1+ meeting or email<5 people)"
*label var weekly_internal_organization_in "# internal departments for which individuals/groups have 1+ interaction (1+ meeting or email<5 people)"
*label var weekly_intra_department_in  "# people that individual/group have interacted with within their own department. (1+ meeting or email<5 people)"
*label var weekly_intra_function_in "# people that individual/group have interacted with within their job function. (1+ meeting or email<5 people)"
*label var weekly_intra_level_interactions # people that individual/group have interacted with within their level. (1+ meeting or email<5 people)
*label var weekly_intra_region_interactions # people that individual/group have interacted with within their location. (1+ meeting or email<5 people)
*label var weekly_leadership_contact # people that individual/group have interacted with above their level. (1+ meeting or email<5 people)
label var timeblock_1hr "# of 1 hour blocks of time between meetings"
label var timeblock_2hrs "# of 2 hour blocks of time between meetings"
label var centrality "Measures how central a person is to the flow of info w/in the company. High centrality means that a person has more connections, and the people that they are connected to also have many connections. "
label var externalnetworkbreadth "# of external domains with whom you have 1+ meaning connection (meeting or email between <5 people.)"
label var externalnetworksize "# distinct external people a person is connected to"
label var internalnetworkbreadth "# internal organizations the person has a meaningful connection with (meeting or email between <5 people.)"
label var internalnetworksize "# distinct external people a person is connected to"
*label var internalnetworkvelocity "# new connections per month added by the members of each group"
*label var intradepartmentalnetworkdepth "# meaningful connections in the time period that the person has had with internal hasGroupings people with the person's WKA organization"
*label var intrafunctionnetworkdepth "# meaningful connections in the time period that the person has had with internal hasGroupings people with the person's WKA function"
*label var intralevelnetworkdepth "# meaningful connections in the time period that the person has had with internal hasGroupings people with the person's WKA level"
*label var intraregionnetworkdepth "# of meaningful connections in the time period that the person has had with internal hasGroupings people with the person's WKA region"
*label var leadershipcontact "Internal network size restricted to internal people with groupings with WKA level above the person"
*label var networkefficiency "= internalCollaborationHoursExtended / internalNetworkSize"
*label var weekly_centrality "Centrality based on interactions with their internal network"
*label var numpeopleinactivity "Number of people with an activity"
*label var numpeoplewithattribute "Number of people with attributes"
*label var timecostinactivity "Time cost with an activity"
*label var timecostwithattribute "Time cost with attributes"
*label var timespentinactivity "Amount of time spent with an activity"
*label var timespentwithattribute "Amount of time spent with attributes"
*label var customereigenvectorcentrality" The person's eigenvector centrality where nodes( internal people & hasGroupings AND external people & isCustomer true)"
*label var customernetworkbreadth "The number of distinct account names from well known attributes that a person is connected to within the 'isCustomer' external set"
*label var customernetworksize "The number of distinct external people that are 'isCustomer' a person is connected to "
label var meetingbreakdownattendees_at "The total time in hours that the person spent in meetings with 21 or more attendees where attendees are non-declined based on adjusted duration"
label var meetingbreakdownattendees_from1 "The total time in hours that the person spent in meetings with 11-20 attendees where attendees are non-declined based on adjusted duration"
label var meetingbreakdownattendees_from3 "The total time in hours that the person spent in meetings with 3-6 attendees where attendees are non-declined based on adjusted duration"
label var meetingbreakdownattendees_from7 "The total time in hours that the person spent in meetings with 7-10 attendees where attendees are non-declined based on adjusted duration"
label var meetingbreakdownattendees_only1 "The total time in hours that the person spent in meetings with 1 attendee where attendees are non-declined based on adjusted duration"
label var meetingbreakdownattendees_only2 "The total time in hours that the person spent in meetings with 2 attendees where attendees are non-declined based on adjusted duration"
*label var timespentwithlevel_above "The time allocated to people in collaboration items above the person's level"
label var timespentwithlevel_below "The time allocated to people in collaboration items below the person's level"
*label var timespentwithlevel_same "The time allocated to people in collaboration items that share the person's level"
label var attendedmeetingquality "A score on a scale of 0 - 100 (where 100 is high) that is used to gauge the average quality of meetings that a pid or group attended using a combination of different meeting metrics"
label var organizedmeetingquality "A score on a scale of 0 - 100 (where 100 is high) that is used to gauge the average quality of meetings that a pid or group organized using a combination of different meeting metrics"

*Redoing theselabels so they appear in the graph in the right hierarchical order*
*Is IC individual contributor? Is this hierarchical order correct?*
replace leveld="1 Director" if leveld=="Director"
replace leveld="2 Senior Executive" if leveld=="Senior Executive"
replace leveld="3 Executive" if leveld=="Executive"
replace leveld="4 Manager" if leveld=="Manager"
replace leveld="5 Senior IC" if leveld=="Senior IC"
replace leveld="6 Junior IC" if leveld=="Junior IC"
replace leveld="7 Support" if leveld=="Support"

*creaing dummy variables*
foreach var in level function region organization {
quietly tab `var', gen(`var')
}

save "KPI_$S_DATE", replace
*/

use "KPI_16 Sep 2016"

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
outreg2 using "regression_basic_$S_DATE", excel label replace
regress attendedmeetingquality level1-level6, cluster(pid)
outreg2 using "regression_basic_$S_DATE", excel label append

foreach var in totalhours meetinghours mailhours overload overloadmeeting overloadmails doublebooked lowengage {
regress `var' function1-function6, cluster(pid)
outreg2 using "regression_basic_$S_DATE", excel label append
regress `var'  level1-level6, cluster(pid)
outreg2 using "regression_basic_$S_DATE", excel label append
}

regress attendedmeetingquality meetinghours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label replace
regress attendedmeetingquality mailhours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality totalhours region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality overloadmeeting region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality overloadmail region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality doublebooked region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality lowenga region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality meetingbreakdownduration* region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append
regress attendedmeetingquality meetingbreakdownatte* region1-region6 function1-function7 level1-level7, cluster(pid)
outreg2 using "regression_quality_$S_DATE", excel label append

regress meetinghours mailhours, cluster(pid)
outreg2 using "regression_$S_DATE", excel label replace

levelsof leveld, local(levels) 
foreach l of local levels {
regress meetinghours mailhours if leveld=="`l'", cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
*regress meetinghours mailhours region2-region6 if leveld=="`l'", cluster(pid)
*outreg2 using "regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
}

levelsof functiont, local(levels) 
foreach l of local levels {
regress meetinghours mailhours if functiont=="`l'", cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
*regress meetinghours mailhours region2-region6 if functiont=="`l'", cluster(pid)
*outreg2 using "regression_hours_$S_DATE", excel label append ctitle(Hours in meetings, "`l'" only)
}

regress meetinghours mailhours level2-level7, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append
regress meetinghours mailhours level2-level7 region2-region6, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append
regress meetinghours mailhours function2-function7, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append
regress meetinghours mailhours function2-function7 region2-region6, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append
regress meetinghours mailhours level2-level7 function2-function7 region2-region6, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append

regress central meetinghours mailhours, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append
regress central meetinghours mailhours level2-level7 function2-function7 region2-region6, cluster(pid)
outreg2 using "regression_hours_$S_DATE", excel label append

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
png2rtf using "microsoft_$S_DATE.doc", g(`var'.png) replace    
erase `var'.png
erase `var'.gph  
}

foreach var in meetinghours meetingsattended attendedmeetingquality  mailhours sentmails receivedmails timeblock_1hr timeblock_2hrs doublebooked utilization internalnetworksize externalnetworksize {
graph bar (mean) `var', over(leveld, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}

foreach var in centrality  meetinghours meetingsattended attendedmeetingquality  mailhours sentmails receivedmails timeblock_1hr timeblock_2hrs doublebooked utilization internalnetworksize externalnetworksize {
graph bar (mean) `var', over(functiont, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "microsoft_$S_DATE.doc", g(`var'.png) append        
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
png2rtf using "microsoft_$S_DATE.doc", g("`l'.png") append        
erase "`l'.png"
erase "`l'.gph"  
}

levelsof functiont, local(levels) 
foreach l of local levels {
graph bar meetingbreakdownduration_lesstha  meetingbreakdownduration_from15 meetingbreakdownduration_from30t meetingbreakdownduration_from1to meetingbreakdownduration_from2to meetingbreakdownduration_atleast if functiont=="`l'", subtitle("Share of Total Meeting Hours by Duration, for `l'") ytitle("Share of Total Meeting Hours") legend(lab(1 "<15min") lab(2 "15-30min") lab(3 "30min - 1hr") lab(4 "1-2 hr") lab(5 "2-8 hr") lab(6 "8hr+") symxsize(5) row(1)) 
graph save "`l'", replace
graph export "`l'.png", replace
png2rtf using "microsoft_$S_DATE.doc", g("`l'.png") append        
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
png2rtf using "microsoft_$S_DATE.doc", g("`l'.png") append        
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
png2rtf using "microsoft_$S_DATE.doc", g("`l'.png") append        
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
png2rtf using "microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}
		
foreach var in share_recurring share_externalmtg share_externalall {
graph bar (mean) `var', over(leveld, label(angle(45))) subtitle("`: variable label `var'', by level", span)
graph save `var', replace
graph export `var'.png, replace
png2rtf using "microsoft_$S_DATE.doc", g(`var'.png) append        
erase `var'.png
erase `var'.gph  
}
		
