clear all



import excel "C:\Users\IBS\Downloads\Output  table_A.xlsx", sheet("Data") firstrow clear

keep entity Marz I1 I2 k2n k3n exp exp_to_02 exp_to_03 exp_to_04 exp_to_05 a0 a1 a4 a6 a8_01 a9_02  a9_03 a9_04 b2 b6a b6b b6c b6d b6e org1_01 org2_00 a9_05 f1_01 f1_02 f1_03 f1_04 f1_05 f1_06 f1_07  f2_01 f2_02 f2_03 f2_04 f2_05 f2_06 f2_07 f2_08

destring Marz, replace force



capture tostring I1, replace force
replace I1 = strtrim(I1)




count
global size_vars "i.k2n i.k3n"

* 1. Create a numeric variable
gen sector_num = .
replace sector_num = 1 if I1 == "C"
replace sector_num = 2 if I1 == "F"
replace sector_num = 3 if I1 == "G"
replace sector_num = 4 if I1 == "H"
replace sector_num = 5 if I1 == "I"
replace sector_num = 6 if I1 == "J"
replace sector_num = 7 if I1 == "K"
replace sector_num = 8 if I1 == "L"
replace sector_num = 9 if I1 == "M"
replace sector_num = 10 if I1 == "N"

* 2. Attach the labels
label define my_sectors1 1 "Manufacturing" 2 "Construction" 3 "Retail/Trade" 4 "Logistics" 5 "Hospitality" 6 "IT Sector" 7 "Finance" 8 "Real Estate" 9 "Consulting" 10 "Administrative and Support Service Activities"
label values sector_num my_sectors

* 1. Clear the old variable to avoid conflicts
capture drop sector_id
capture label drop my_sectors

* 2. Re-create sector_id from I1 (Alphabetical Order)
* Stata sees: C, D, E, F, G, H, I, J, L, M, N, S
encode I1, gen(sector_id)

label define my_sectors ///
    1 "C: Manufacturing" ///
    2 "D: Electricity/Gas" ///
    3 "E: Water/Waste" ///
    4 "F: Construction" ///
    5 "G: Trade (Retail)" ///
    6 "H: Transport" ///
    7 "I: Hospitality" ///
    8 "J: IT Sector" ///
    9 "L: Real Estate" ///
    10 "M: Consulting/Science" ///
    11 "N: Admin/Support" ///
    12 "S: Other Services", replace

label values sector_id my_sectors


tab sector_id

tab sector_num

replace a0 = "1" if regexm(a0, "Yes")
replace a0 = "0" if regexm(a0, "No")
destring a0, replace force

replace a4 = "1" if regexm(a4, "Yes")
replace a4 = "0" if regexm(a4, "No")
destring a4, replace force

replace a6 = "1" if regexm(a6, "Yes")
replace a6 = "0" if regexm(a6, "No")
destring a6, replace force

replace org1_01 = "1" if regexm(org1_01, "Yes")
replace org1_01 = "0" if regexm(org1_01, "No")
destring org1_01, replace force

destring org2_00, replace force
recode org2_00 (0=1) (1=0) (.=0)
label variable org2_00 "CRM Software (1=Yes)"

recode a0 (2=0) (.=0) (0=0) (1=1)
recode a1 (2=0) (.=0) (0=0) (1=1)
recode a4 (2=0) (.=0) (0=0) (1=1)
recode a6 (2=0) (.=0) (0=0) (1=1)
recode a8_01 (2=0) (.=0) (0=0) (1=1)
recode org1_01 (2=0) (.=0) (0=0) (1=1)
recode a9_02 (2=0) (.=0) (0=0) (1=1)
recode a9_03 (2=0) (.=0) (0=0) (1=1)
recode a9_04 (2=0) (.=0) (0=0) (1=1)
recode a9_05 (2=0) (.=0) (0=0) (1=1)

gen w_score = 0
replace w_score = w_score + 1 if a0 == 1
replace w_score = w_score + 1 if a4 == 1

* TIER 2 (2 Points)
replace w_score = w_score + 2 if a6 == 1
replace w_score = w_score + 2 if a8_01 == 1

* TIER 3 (3 Points)
replace w_score = w_score + 3 if org1_01 == 1
replace w_score = w_score + 3 if org2_00 == 1
replace w_score = w_score + 3 if a9_02 == 1
replace w_score = w_score + 3 if a9_03 == 1

* TIER 4 (5 Points)
replace w_score = w_score + 5 if a9_04 == 1
summarize w_score, detail

drop if w_score == 0

recode exp_to_02 exp_to_03 exp_to_04 exp_to_05 (.=0)
recode b6a b6b b6c b6d b6e (.=0)

recode exp b2 (.=0)

recode a0 a1 a4 a6 a8_01 a9_02 a9_03 a9_04 a9_05 org1_01 org2_00 (.=0)
recode f1_01 f1_02 f1_03 f1_04 f1_05 f1_06 f1_07 (.=0)
recode f2_01 f2_02 f2_03 f2_04 f2_05 f2_06 f2_07 f2_08 (.=0)

destring k2n k3n, replace force
gen d_exp_eeu = (exp_to_02 == 1) /* Russia/Belarus */
gen d_exp_cis = (exp_to_03 == 1) /* Other CIS */
gen d_exp_eu  = (exp_to_04 == 1) /* Europe */
gen d_exp_oth = (exp_to_05 == 1) /* USA/Rest of World */

gen d_online_arm = (b6a > 0)
gen d_online_eeu = (b6b > 0)
gen d_online_cis = (b6c > 0)
gen d_online_eu  = (b6d > 0)
gen d_online_oth = (b6e > 0)

gen op_ai = 0
replace op_ai = 1 if f2_02 == 1 | f2_05 == 1

gen a9_05_pct = a9_05 * 100
label variable a9_05_pct "GenAI Adoption (%)"

* 2. Table 1: Adoption Percentage by Sector

tabulate sector_id, summarize(a9_05_pct)

* 3. Table 2: AI Components Among Users
gen is_ai_user = (a9_05 == 1 | op_ai == 1)

* Create percentage versions for the components too
gen op_ai_pct = op_ai * 100
gen cloud_pct = a9_03 * 100
gen crm_pct   = org2_00 * 100 /* Assuming 1=Used */

* Show the table restricted to users only
tabstat a9_05_pct op_ai_pct cloud_pct crm_pct if is_ai_user == 1, ///
    statistics(mean count) columns(statistics)
	
* =========================================================
* FINAL REGRESSION TABLE: LOGIT MARGINAL EFFECTS
* =========================================================

* Install the export tool if you haven't yet
ssc install estout, replace

* ---------------------------------------------------------
* MODEL 1: BASELINE (GenAI, Few Controls)
* ---------------------------------------------------------
* We run Logit, then calculate Margins, then 'post' the results so esttab can see them.
logit a9_05 d_online_eu d_exp_eu c.w_score i.k2n, vce(robust)
margins, dydx(*) post
eststo model1

* ---------------------------------------------------------
* MODEL 2: FULL SPECIFICATION (GenAI, All Controls)
* ---------------------------------------------------------
logit a9_05 d_exp_eeu d_exp_cis d_exp_eu d_exp_oth d_online_eeu d_online_cis d_online_eu d_online_oth c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model2

* ---------------------------------------------------------
* MODEL 3: PLACEBO TEST (Operational AI)
* ---------------------------------------------------------
logit op_ai d_exp_eeu d_exp_cis d_exp_eu d_exp_oth d_online_eeu d_online_cis d_online_eu d_online_oth c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model3

* ---------------------------------------------------------
* EXPORT TO LATEX
* ---------------------------------------------------------
esttab model1 model2 model3 using "logit_margins.tex", replace ///
    b(3) se(3) star(* 0.10 ** 0.05 *** 0.01) ///
    label booktabs nonotes ///
    title("Determinants of AI Adoption: Average Marginal Effects") ///
    mtitle("Gen AI (Base)" "Gen AI (Full)" "Ops AI (Placebo)") ///
    keep(d_online_eu d_exp_eu w_score d_online_eeu) ///
    scalars("N Observations") ///
    addnote("Values represent Average Marginal Effects (dy/dx). Standard errors in parentheses.")
	
	* =========================================================
* DESCRIPTIVE STATISTICS GENERATOR
* =========================================================

* 1. PREPARE VARIABLES (Percent 0-100)
gen gen_ai_pct = a9_05 * 100
label var gen_ai_pct "Generative AI (%)"

* Administrative AI (The Purpose - f2_03)
* Ensure f2_03 is clean (2=0, .=0)
capture confirm string variable f2_03
if _rc == 0 {
    replace f2_03 = "1" if regexm(f2_03, "[Yy]es")
    replace f2_03 = "0" if regexm(f2_03, "[Nn]o")
    destring f2_03, replace force
}
recode f2_03 (2=0) (.=0) (0=0) (1=1)
gen admin_ai_pct = f2_03 * 100

* Production AI (The Purpose - f2_02)
gen prod_ai_pct = f2_02 * 100

* Marketing AI (The Purpose - f2_01)
gen mkt_ai_pct = f2_01 * 100

* ---------------------------------------------------------
* TABLE 1: ADOPTION BY SECTOR
* ---------------------------------------------------------
display " "
display "--- TABLE 1: ADOPTION RATES BY SECTOR (%) ---"
tabulate sector_id, summarize(gen_ai_pct)
tabulate sector_id, summarize(admin_ai_pct)

graph bar (mean) gen_ai_pct, over(sector_id, label(angle(45) labsize(small))) ///
    title("Generative AI Adoption by Sector") ///
    ytitle("Adoption Rate (%)") ///
    bar(1, color(navy)) ///
    blabel(bar, format(%9.1f)) ///
    scheme(s1color)

graph export "sector_adoption.png", replace

preserve
    keep if a9_05 == 1
    gen admin_pct = f2_03 * 100
    gen prod_pct  = f2_02 * 100
    gen mkt_pct   = f2_01 * 100
    
    collapse (mean) admin_pct prod_pct mkt_pct
    gen id = 1
    reshape long @_pct, i(id) j(tech_type) string
    
    replace tech_type = "Administrative" if tech_type == "admin"
    replace tech_type = "Production" if tech_type == "prod"
    replace tech_type = "Marketing" if tech_type == "mkt"

    graph hbar (asis) _pct, over(tech_type, sort(1) descending) ///
        title("Types of AI Usage Among GenAI Adopters") ///
        ytitle("Percentage of Adopters") ///
        ylabel(0(20)100) ///
        blabel(bar, format(%9.1f)) ///
        bar(1, color(maroon)) ///
        scheme(s1color)

    graph export "ai_usage.png", replace
restore
quietly regress a9_05 d_online_eu d_exp_eu w_score k2n sector_id
vif
quietly regress a9_05 d_online_eu d_exp_eu w_score k2n
estat hettest
logit a9_05 d_online_eu d_exp_eu w_score i.k2n i.sector_id, vce(robust)
lroc
* =========================================================
* EXPORT FULL TABLES (FOR APPENDIX/SLIDES)
* =========================================================

quietly logit a9_05 d_online_eu d_exp_eu c.w_score i.k2n, vce(robust)
margins, dydx(*) post
eststo model1

quietly logit a9_05 d_exp_eeu d_exp_cis d_exp_eu d_exp_oth ///
            d_online_eeu d_online_cis d_online_eu d_online_oth ///
            c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model2

quietly logit op_ai d_exp_eeu d_exp_cis d_exp_eu d_exp_oth ///
            d_online_eeu d_online_cis d_online_eu d_online_oth ///
            c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model3

* 2. EXPORT EVERYTHING (No 'keep' command)
* We use 'drop' just to hide the 'base' empty categories if needed, 
* but here we just let it show everything.
esttab model1 model2 model3 using "full_appendix_table.tex", replace ///
    b(3) se(3) star(* 0.10 ** 0.05 *** 0.01) ///
    label booktabs nonotes ///
    title("Appendix: Full Regression Output") ///
    mtitle("Gen AI (Base)" "Gen AI (Full)" "Ops AI (Placebo)") ///
    scalars("N Observations") ///
    addnote("Values are Average Marginal Effects.")
	

preserve
    * 1. Keep ONLY Generative AI Users
    keep if a9_05 == 1

    * 2. Clean ALL f2 variables (Purposes)
    local purposes f2_01 f2_02 f2_03 f2_04 f2_05 f2_06 f2_07 f2_08
    foreach v of local purposes {
        capture confirm string variable `v'
        if _rc == 0 {
            replace `v' = "1" if regexm(`v', "[Yy]es")
            replace `v' = "0" if regexm(`v', "[Nn]o")
            destring `v', replace force
        }
        recode `v' (2=0) (.=0) (0=0) (1=1)
        replace `v' = `v' * 100
    }

    * 3. Collapse to get the average usage rates
    collapse (mean) f2_01 f2_02 f2_03 f2_04 f2_05 f2_06 f2_07 f2_08

    * 4. Reshape for Graphing
    gen id = 1
    reshape long f2_, i(id) j(purpose_code) string

    * 5. Apply Labels
    gen purpose_label = ""
    replace purpose_label = "Marketing & Sales" if purpose_code == "01"
    replace purpose_label = "Production Processes" if purpose_code == "02"
    replace purpose_label = "Administration" if purpose_code == "03"
    replace purpose_label = "Product Design" if purpose_code == "04"
    replace purpose_label = "Logistics" if purpose_code == "05"
    replace purpose_label = "Security" if purpose_code == "06"
    replace purpose_label = "HR & Recruiting" if purpose_code == "07"
    replace purpose_label = "Finance" if purpose_code == "08"
    replace purpose_label = "Other" if purpose_label == ""

    * 6. Generate the Graph (FIXED: grid is inside ylabel)
    graph hbar (asis) f2_, over(purpose_label, sort(1) descending label(labsize(small))) ///
        title("For Which Purposes Do Firms Use AI?") ///
        subtitle("Sample: Firms adopting Generative AI") ///
        ytitle("Percentage of Adopters") ///
        ylabel(0(10)60, grid) ///   <--- FIXED HERE
        blabel(bar, format(%9.1f) size(small)) ///
        bar(1, color(steelblue)) ///
        scheme(s1color)

    * 7. Save
    graph export "ai_purposes_graph.png", replace
	
restore

* =========================================================
* ROBUSTNESS CHECK: EXCLUDING THE IT SECTOR
* =========================================================


logit a9_05 d_online_eu d_exp_eu c.w_score i.k2n i.sector_id if sector_id != 8, vce(robust)
margins, dydx(*)

* =========================================================
* FINAL REGRESSION EXPORT (FULL TABLE)
* =========================================================

* 1. Run Model 1: Generative AI (Baseline)
quietly logit a9_05 d_online_eu d_exp_eu c.w_score i.k2n, vce(robust)
margins, dydx(*) post
eststo model1

* 2. Run Model 2: Generative AI (Full Specification - MAIN RESULT)
quietly logit a9_05 d_exp_eeu d_exp_cis d_exp_eu d_exp_oth ///
            d_online_eeu d_online_cis d_online_eu d_online_oth ///
            c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model2

* 3. Run Model 3: Operational AI (Placebo Test)
quietly logit op_ai d_exp_eeu d_exp_cis d_exp_eu d_exp_oth ///
            d_online_eeu d_online_cis d_online_eu d_online_oth ///
            c.w_score i.k3n i.k2n i.sector_id, vce(robust)
margins, dydx(*) post
eststo model3

* 4. Export to LaTeX File
esttab model1 model2 model3 using "full_regression_results.tex", replace ///
    b(3) se(3) star(* 0.10 ** 0.05 *** 0.01) ///
    label booktabs nonotes ///
    title("Determinants of AI Adoption: Full Results") ///
    mtitle("Gen AI (Base)" "Gen AI (Full)" "Ops AI (Placebo)") ///
    scalars("N Observations") ///
    addnote("Values represent Average Marginal Effects (dy/dx). Standard errors in parentheses.")
