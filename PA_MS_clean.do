/**********************************************************************************File        : PA_MS_code.do  <--- this should be the exact name of THIS documentAuthor      : Kristina Tobio Created     : 23 Feb 2017Modified    : 12 Sep 2017Description : .do file for MS case**********************************************************************************/capture log closeclearset more offlog using "logs/ms_$S_DATE.log", replaceuse "data/PA_MS_data"/*replicating tables in the case*/tab functionsum totalcount emailcount  meetingcountsum totalcount_int emailcount_int meetingcount_intsum totalcount_ext emailcount_ext meetingcount_extkeep if function=="Sales"sum totalcount emailcount  meetingcountsum totalcount_int emailcount_int meetingcount_intsum totalcount_ext emailcount_ext meetingcount_extclear/*LOGS*/use "data/PA_MS_data"corr totalcount attainmentcorr emailcount attainmentcorr meetingcount attainmentg ltotalcount=log(totalcount)g lattainment=log(attainment)regress attainment totalcountregress attainment ltotalcountregress lattainment ltotalregress attainment totalcount leveldum2-leveldum7 regiondum2-regiondum6 outreg2 using "reports/logs_$S_DATE", replace bdec(3) pvalue label excelregress attainment ltotalcount leveldum2-leveldum7 regiondum2-regiondum6 outreg2 using "reports/logs_$S_DATE", append bdec(3) pvalue label excelregress lattainment ltotal leveldum2-leveldum7 regiondum2-regiondum6 outreg2 using "reports/logs_$S_DATE", append bdec(3) pvalue label excelclearuse "data/PA_MS_data", replacetab functiondkeep if function=="Sales"hist totalcount, freqgraph save "reports/totalcount", replacegraph export "reports/totalcount.png", replacepng2rtf using "reports/hist.doc", g("reports/totalcount.png") replace    erase "reports/totalcount.png"hist emailcount, freqgraph save "reports/emailcount"graph export "reports/emailcount.png", replacepng2rtf using "reports/hist.doc", g("reports/emailcount.png") append   erase "reports/emailcount.png"hist meetingcount, freqgraph save "reports/meetingcount", replacegraph export "reports/meetingcount.png", replacepng2rtf using "reports/hist.doc", g("reports/meetingcount.png") append    erase "reports/meetingcount.png"hist totalcount_ext, freqgraph save "reports/totalcount", replacegraph export "reports/totalcount_ext.png", replacepng2rtf using "reports/hist.doc", g("reports/totalcount_ext.png") append    erase "reports/totalcount_ext.png"hist emailcount_ext, freqgraph save "reports/emailcount", replacegraph export "reports/emailcount_ext.png", replacepng2rtf using "reports/hist.doc", g("reports/emailcount_ext.png") append    erase "reports/emailcount_ext.png"hist meetingcount_ext, freqgraph save "reports/meetingcount", replacegraph export "reports/meetingcount_ext.png", replacepng2rtf using "reports/hist.doc", g("reports/meetingcount_ext.png") append    erase "reports/meetingcount_ext.png"hist totalcount_int, freqgraph save "reports/totalcount", replacegraph export "reports/totalcount_int.png", replacepng2rtf using "reports/hist.doc", g("reports/totalcount_int.png") append    erase "reports/totalcount_int.png"hist emailcount_int, freqgraph save "reports/emailcount", replacegraph export "reports/emailcount_int.png", replacepng2rtf using "reports/hist.doc", g("reports/emailcount_int.png") append    erase "reports/emailcount_int.png"hist meetingcount_int, freqgraph save "reports/meetingcount", replacegraph export "reports/meetingcount_int.png", replacepng2rtf using "reports/hist.doc", g("reports/meetingcount_int.png") append    erase "reports/meetingcount_int.png"clear*drop if totalcount<1000*drop if totalcount>20000use "data/PA_MS_data"regress attain  totalcountoutreg2 using "reports/MS_output_$S_DATE", replace bdec(3) pvalue label excel addnote("Missing indicators are level=Director; region=Central; Group=C")regress attain  emailcountoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcountoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_extoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_extoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_extoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_ext totalcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_ext emailcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_ext meetingcount_intoutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label exceloutreg2 using "reports/MS_output_$S_DATE", replace bdec(3) pvalue label excel addnote("Missing indicators are level=Director; region=Central; Group=C")regress attain  emailcount leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_ext leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_ext leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_ext leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_int leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_int leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_int leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  totalcount_ext totalcount_int  leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  emailcount_ext emailcount_int  leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  meetingcount_ext meetingcount_int  leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *total totalcount_ext outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *email emailcount_ext outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *mtg meetingcount_ext outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *total totalcount_ext  leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *email  emailcount_ext leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain  *mtg meetingcount_ext leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excel rename level_num numforeach y in total mtg email {foreach x in PM HR Mkt Op GA Sales RD {regress attain num  `x'`y' outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain num  `x'`y' leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excel}}corr attain num centralregress attain num centraloutreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excelregress attain num  central leveldum2-leveldum7 regiondum1-regiondum6  outreg2 using "reports/MS_output_$S_DATE", append bdec(3) pvalue label excel