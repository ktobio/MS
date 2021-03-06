clear
clear matrix
capture log close
set more off

log using "logs/MS_cleaning.log", replace

*CLEANING THE DATA AND PREPARING FOR ANALYSIS*

*insheeting .csv. Renamed "Demo KPI Dump - All Dates_v2.csv" to "KPI.csv", no other changes to file were made*

insheet using "./data/Demo KPI Dump - All Dates_v2.csv", names case

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

save "data/KPI", replace

