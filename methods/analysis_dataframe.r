# This is code to analse the data collected for the systematic literature review
# Code developed by David Pedrosa

# Version 1.3 # 2022-03-19, added analyses for QES

# ==================================================================================================
## Specify packages of interest and load them automatically if needed
packages = c("readxl", "dplyr", "plyr", "tibble", "tidyr", "stringr", "openxlsx", 
				"metafor", "tidyverse") 											# packages needed

## Load or install necessary packages
package.check <- lapply(
  packages,
  FUN = function(x) {
    if (!require(x, character.only = TRUE)) {
      install.packages(x, dependencies = TRUE)
      library(x, character.only = TRUE)
    }
  }
)

# ==================================================================================================
## In case of multiple people working on one project, this helps to create an automatic script
username = Sys.info()["login"]
if (username == "dpedr") {
wdir = "D:/Jottacloud/PDtremor_syst-review/"
} else if (username == "david") {
wdir = "D:/Jottacloud/PDtremor_syst-review/"
}
setwd(wdir)

# ==================================================================================================
### PART I: GENERAL DATA EXTRACTION FROM WOKSHEETS
# ==================================================================================================
# 	Structure of csv-file: author	|	year	|	title	|	journal	|	group	|	study_design	|	objective	|	n	|	primary_outcomes	|	time_to_follow_up

results_file 	<-  file.path(wdir, "__included_studies.xlsx")
excel_sheets(path = results_file) 								# Names of sheets for later use
tab_names 		<- excel_sheets(path = results_file)

# ==================================================================================================
# Extract authors and years from worksheets
authors 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('Abbr', 'Author'), range = "C2:D2"))

years 		<- lapply(tab_names, function(x) read_excel(results_file, 															
															sheet = x, col_names=c('Abbr', 'Year'), range = "C4:D4"))
authors_wide<- data.frame(data.table::rbindlist(authors, idcol='ID')) %>% select(Author) # data.frame(Author=sapply(Authors, "[[", 2))
years_short <- data.frame(data.table::rbindlist(years, idcol='ID')) %>% select(Year) # data.frame(years=sapply(years, "[[", 2))

# Clean up author list so that it can be incorporated to csv table as xxx et al. 19xx/20xx
# test regexp code here: https://regexr.com/ or https://spannbaueradam.shinyapps.io/r_regex_tester/
pattern 		<- "([\\w\\-\\,\\s]+)[ ]+([A-Za-z])*.([\\.\\,\\s]*)"# pattern to extract authors as xxx et al. (cf. line 62f. for exception)
authors_short 	<- authors_wide$Author %>% str_extract_all(pattern) # creates list of authors with 1 or maximum 2 authors according to APA 
authors_short 	<- sapply(authors_short, "[[", 1)

pattern 		<- "([\\w\\-\\,\\s]+)[ ]+([A-Za-z])*.([\\.\\,\\s]*)"
authors_total 	<- authors_wide$Author %>% str_extract_all(pattern) # creates list of authors with 1 or maximum 2 authors according to APA recommendations 
idx_authors_man <- which(sapply(authors_total, length)<=2)	# index of studies requiring manual changes (e.g., only two authors or within "study groups")
authors_short 	<- sapply(authors_total, "[[", 1) # get a list of all first authors
first_authors 	<- paste(sapply( str_split(authors_short, ','), "[[",1), "et al.", years_short$Year)

# Manual changes for studies with < 3 authors 
for (i in idx_authors_man){
	author1 = str_split(authors_total[[i]][1], ',')
	if (length(authors_total[[i]])==1) {
		first_authors[i] = paste0(author1[[1]][1], " (", years_short$Year[i] ,")")
	} else {
		author2 = str_split(authors_total[[i]][2], ',')
		first_authors[i] = paste(author1[[1]][1], "&", str_squish(author2[[1]][1]), paste0("(", years_short$Year[i] ,")"))
	}
}
first_authors <- make.unique(first_authors)

# ==================================================================================================
# Extract title, journal and year from worksheets
extracted_data1	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C4:D6"))
extracted_data1_wide = extracted_data1 %>% bind_rows(extracted_data1, .id="sheet") %>% 
  dplyr::rename(category = 2, value = 3) %>% 
  pivot_wider(names_from = category, values_from = value) %>%
  filter(as.numeric(sheet) <= length(first_authors))
  
# ==================================================================================================
# Extract group from worksheets, that is category such as dopamine_agonists, levodopa, etc.
extracted_group 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C1:D1"))
groups_short 		<- data.frame(substance=sapply(extracted_group, "[[", 2))
nmax <- max(stringr::str_count(groups_short$substance, ";")) + 1 # maximum number of groups;
groups_short = tidyr::separate(groups_short, substance, paste0("intervention", seq_len(nmax)), sep = ";", fill = "right") # create one colum per intervention

# ==================================================================================================
# Extract study type from worksheets
extracted_study_type 		<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C10:D10"))
study_type_short 	<- data.frame(study_design=sapply(extracted_study_type, "[[", 2)) # get a list of all study types encountered

# ==================================================================================================
# Extract intervention from worksheets
extracted_intervention 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('abbr', 'intervention'), range = "C33:D33"))
intervention_short 		<- data.frame(description=sapply(extracted_intervention, "[[", 2)) # get a list of all study types encountered

# ==================================================================================================
# Extract aims from worksheets
extracted_aims 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C9:D9"))
aims_short 		<- data.frame(aims=sapply(extracted_aims, "[[", 2)) # get a list of all study types encountered

# ==================================================================================================
# Extract inclusion/exclusion criteria from worksheets
extracted_criteria	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C18:D18"))
crit1				<- data.frame(inclusion=sapply(extracted_criteria, "[[", 2)) # get a dataframe with all criteria
nmax 				<- max(stringr::str_count(crit1$inclusion, ";"), na.rm=TRUE) + 1 # maximum number of groups;
inclusion_short 	<- tidyr::separate(crit1, inclusion, paste0("inclusion_criterion", seq_len(nmax)), sep = ";", fill = "right") # create one colum per intervention

extracted_criteria	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=FALSE, range = "C19:D19"))
crit2				<- data.frame(exclusion=sapply(extracted_criteria, "[[", 2)) # get a dataframe with all criteria
nmax 				<- max(stringr::str_count(crit2$exclusion, ";"), na.rm=TRUE) + 1 # maximum number of groups;
exclusion_short 	<- tidyr::separate(crit2, exclusion, paste0("exclusion_criterion", seq_len(nmax)), sep = ";", fill = "right") # create one colum per intervention

total <- cbind(extracted_data1_wide, "Authors" = first_authors)
total <- cbind(total, groups_short, intervention_short, study_type_short, aims_short, inclusion_short, exclusion_short)

total <- total %>% select(Authors, Title, Journal, Year, study_design, aims, intervention1,
							intervention2, intervention3, description, colnames(inclusion_short), colnames(exclusion_short))
							
total %>% write.csv("synopsis.PDtremorSystematicReview.csv", row.names = F)

total %>% write.xlsx("synopsis.PDtremorSystematicReview.xlsx", sheetName = "Synopsis_results", 
  col.names = TRUE, row.names = FALSE, append = FALSE, overwrite=TRUE)


# ==================================================================================================
### PART II: EXTRACTION OF PRIMARY AND SECONDARY OUTCOMES FROM WORKSHEETS
# ==================================================================================================
# Extract primary outcomes
extracted_outcomes_p 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome'), range = "C37:C44"))

all_outcomes_p 			<- data.frame(data.table::rbindlist(extracted_outcomes_p, idcol='ID')) # merges everything to one dataframe

outcomes_wide_p 		<- data.frame(matrix(ncol=10, nrow=length(authors_short)))
outcomes_wide_p[,1]		<- first_authors
outcomes_wide_p[,2]		<- study_type_short
colnames(outcomes_wide_p) <- c('author', 'study_type', paste0('primary_outcome', c(1:8)))

for (i in unique(all_outcomes_p$ID)) {
	idx = which(all_outcomes_p==i)
	outcomes_wide_p[i,3:10] <- all_outcomes_p[which(all_outcomes_p==i),]$outcome %>% t
}

outcomes_wideRCTprim <- outcomes_wide_p[which(outcomes_wide_p$study_type=="Randomised-controlled trial (RCT)"),]
outcomes_wideRCTprim %>% write.xlsx("primary_outcomesRCT.xlsx", sheetName = "RCT", 
  col.names = TRUE, row.names = TRUE, append = FALSE, overwrite=TRUE) # This part merely serves to identify studies with insufficient/wrong aligned information

outcomes_wideQESprim <- outcomes_wide_p[which(outcomes_wide_p$study_type=="Quasi-experimental study"),]
outcomes_wideQESprim %>% write.xlsx("primary_outcomesQES.xlsx", sheetName = "QES", 
  col.names = TRUE, row.names = TRUE, append = FALSE, overwrite=TRUE) # This part merely serves to identify studies with insufficient/wrong aligned information

# Extract secondary outcomes
extracted_outcomes_s 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome'), range = "C46:C64"))

all_outcomes_s 			<- data.frame(data.table::rbindlist(extracted_outcomes_s, idcol='ID')) # merges everything to one dataframe

outcomes_wide_s 		<- data.frame(matrix(ncol=21, nrow=length(authors_short)))
outcomes_wide_s[,1]		<- first_authors
outcomes_wide_s[,2]		<- study_type_short
colnames(outcomes_wide_s) <- c('author', 'study_type', paste0('secondary_outcome', c(1:19)))

for (i in unique(all_outcomes_s$ID)) {
	idx = which(all_outcomes_s==i)
	outcomes_wide_s[i,3:21] <- all_outcomes_s[which(all_outcomes_s==i),]$outcome %>% t
}

selection <- outcomes_wideRCTprim$author[which(is.na(outcomes_wideRCTprim$primary_outcome1))] # select studies for which there was no primary outcome
outcomes_wideRCTsec <- outcomes_wide_s %>% filter(study_type %in% "Randomised-controlled trial (RCT)") %>% filter(author %in% selection)

outcomes_wideRCTsec %>% write.xlsx("secondary_outcomesRCT.xlsx", sheetName = "RCT", 
  col.names = TRUE, row.names = TRUE, append = FALSE, overwrite=TRUE) # This part merely serves to identify studies with insufficient/wrong aligned information

outcomes_wideQESsec <- outcomes_wide_s %>% filter(study_type %in% "Quasi-experimental study") %>% filter(author %in% selection)

outcomes_wideQESsec %>% write.xlsx("secondary_outcomesQES.xlsx", sheetName = "RCT", 
  col.names = TRUE, row.names = TRUE, append = FALSE, overwrite=TRUE) # This part merely serves to identify studies with insufficient/wrong aligned information


# ==================================================================================================
### PART III: CATEGORIZE AVAILABLE DATA INTO SIMILAR GROUPS AND SYNTHESIZE WITH EFFECT SIZES
# ==================================================================================================

# 1. Extract all relevant primary and secondary outcomes from RCTs with baseline and follow-up values (blfu)
# A. Primary outcomes	
outcomes_rct.blfu.prim 		<- outcomes_wide_p %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Randomised-controlled trial (RCT)") %>%
								filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .)))
	
data_temp1 					<- outcomes_rct.blfu.prim %>% select(!contains("outcome"))
data_temp2 					<- outcomes_rct.blfu.prim %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:8))
outcomes_rct.blfu.prim 		<- cbind(data_temp1, data_temp2)

tibble_rct.blfu.prim 		<- lapply(tab_names[as.numeric(outcomes_rct.blfu.prim$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C37:J44"))
primary_outcomes.blfu 		<- data.frame(data.table::rbindlist(tibble_rct.blfu.prim, idcol='ID')) # merges everything to one dataframe


# B. Secondary outcomes
outcomes_rct.blfu.sec 		<- outcomes_wide_s %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Randomised-controlled trial (RCT)") %>%
								filter(across(secondary_outcome1, ~ grepl('baseline|follow-up', .)))
	
data_temp1 					<- outcomes_rct.blfu.sec %>% select(!contains("outcome"))
data_temp2 					<- outcomes_rct.blfu.sec %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:19))
outcomes_rct.blfu.sec 		<- cbind(data_temp1, data_temp2)


tibble_rct.blfu.sec 		<- lapply(tab_names[as.numeric(outcomes_rct.blfu.sec$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C45:J64"))
secondary_outcomes.blfu 	<- data.frame(data.table::rbindlist(tibble_rct.blfu.sec, idcol='ID')) %>% drop_na(outcome) # merges everything to one dataframe

# Merge results for baseline + follow-up data (blfu)
tibble_rct.blfu 			<- list(tibble_rct.blfu.prim, tibble_rct.blfu.sec) 		# facilitates data handling in a loop later
outcomes_rct.all.blfu 		<- list(primary_outcomes.blfu, secondary_outcomes.blfu) # facilitates data handling in a loop later ## NOT USED, REMOVE?
outcomes_rct.blfu 			<- list(outcomes_rct.blfu.prim, outcomes_rct.blfu.sec)

# ==================================================================================================
# 2. Extract all relevant primary and secondary outcomes from QES with baseline and follow-up values (blfu)
# A. Primary outcomes	

outcomes_qes.blfu.prim 		<- outcomes_wide_p %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Quasi-experimental study") %>%
								filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .)))
	
data_temp1 					<- outcomes_qes.blfu.prim %>% select(!contains("outcome"))
data_temp2 					<- outcomes_qes.blfu.prim %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:8))
outcomes_qes.blfu.prim 		<- cbind(data_temp1, data_temp2)

tibble_qes.blfu.prim 		<- lapply(tab_names[as.numeric(outcomes_qes.blfu.prim$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C37:J44"))
primary_outcomes.blfu 		<- data.frame(data.table::rbindlist(tibble_qes.blfu.prim, idcol='ID')) # merges everything to one dataframe


# B. Secondary outcomes
outcomes_qes.blfu.sec 		<- outcomes_wide_s %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Quasi-experimental study") %>%
								filter(across(secondary_outcome1, ~ grepl('baseline|follow-up', .)))
	
data_temp1 					<- outcomes_qes.blfu.sec %>% select(!contains("outcome"))
data_temp2 					<- outcomes_qes.blfu.sec %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:19))
outcomes_qes.blfu.sec 		<- cbind(data_temp1, data_temp2)


tibble_qes.blfu.sec 		<- lapply(tab_names[as.numeric(outcomes_qes.blfu.sec$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C45:J64"))
secondary_outcomes.blfu 	<- data.frame(data.table::rbindlist(tibble_qes.blfu.sec, idcol='ID')) %>% drop_na(outcome) # merges everything to one dataframe

# Merge results for baseline + follow-up data (blfu)
tibble_qes.blfu 			<- list(tibble_qes.blfu.prim, tibble_qes.blfu.sec) 		# facilitates data handling in a loop later
outcomes_qes.all.blfu 		<- list(primary_outcomes.blfu, secondary_outcomes.blfu) # facilitates data handling in a loop later ## NOT USED, REMOVE?
outcomes_qes.blfu 			<- list(outcomes_qes.blfu.prim, outcomes_qes.blfu.sec)

# ==================================================================================================
# 3. Extract relevant primary and secondary outcomes from RCTs with results as mean-difference (md)
# A. Primary outcomes
outcomes_rct.md.prim 		<- outcomes_wide_p %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Randomised-controlled trial (RCT)" ) %>%
								filter(if_any(everything(), ~ grepl("mean-difference", .)))

data_temp1 					<- outcomes_rct.md.prim %>% select(!contains("outcome"))
data_temp2 					<- outcomes_rct.md.prim %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:8))
outcomes_rct.md.prim 		<- cbind(data_temp1, data_temp2)

tibble_rct.md.prim 			<- lapply(tab_names[as.numeric(outcomes_rct.md.prim$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C37:J44"))
primary_outcomes.md 		<- data.frame(data.table::rbindlist(tibble_rct.md.prim, idcol='ID')) # merges everything to one dataframe

# B. Secondary outcomes
outcomes_rct.md.sec 		<- outcomes_wide_s %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Randomised-controlled trial (RCT)" ) %>%
								filter(if_any(everything(), ~ grepl("mean-difference", .)))

data_temp1 					<- outcomes_rct.md.sec %>% select(!contains("outcome"))
data_temp2 					<- outcomes_rct.md.sec %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:19))
outcomes_rct.md.sec 		<- cbind(data_temp1, data_temp2)

tibble_rct.md.sec 			<- lapply(tab_names[as.numeric(outcomes_rct.md.sec$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C45:J64"))
secondary_outcomes.md 		<- data.frame(data.table::rbindlist(tibble_rct.md.sec, idcol='ID')) # merges everything to one dataframe

# Merge results for mean-difference data (md)
tibble_rct.md 				<- list(tibble_rct.md.prim, tibble_rct.md.sec) 		# facilitates data handling in a loop later
outcomes_all.md 			<- list(primary_outcomes.md, secondary_outcomes.md) # facilitates data handling in a loop later ## NOT USED, REMOVE?
outcomes_rct.md 			<- list(outcomes_rct.md.prim, outcomes_rct.md.sec)

# ==================================================================================================
# 4. Extract relevant primary and secondary outcomes from QES with results as mean-difference (md)
# A. Primary outcomes
outcomes_qes.md.prim 		<- outcomes_wide_p %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Quasi-experimental study" ) %>%
								filter(if_any(everything(), ~ grepl("mean-difference", .)))

data_temp1 					<- outcomes_qes.md.prim %>% select(!contains("outcome"))
data_temp2 					<- outcomes_qes.md.prim %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:8))
outcomes_qes.md.prim 		<- cbind(data_temp1, data_temp2)

tibble_qes.md.prim 			<- lapply(tab_names[as.numeric(outcomes_qes.md.prim$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C37:J44"))
primary_outcomes.md 		<- data.frame(data.table::rbindlist(tibble_qes.md.prim, idcol='ID')) # merges everything to one dataframe

# B. Secondary outcomes
outcomes_qes.md.sec 		<- outcomes_wide_s %>% # somehow bulky way that MUST be simplified somehow using e.g., a lapply function 
								rownames_to_column("type") %>%
								filter(study_type=="Quasi-experimental study" ) %>%
								filter(if_any(everything(), ~ grepl("mean-difference", .)))

data_temp1 					<- outcomes_qes.md.sec %>% select(!contains("outcome"))
data_temp2 					<- outcomes_qes.md.sec %>% 	
								select(contains("outcome")) %>% set_names(paste0("outcome", 1:19))
outcomes_qes.md.sec 		<- cbind(data_temp1, data_temp2)

tibble_qes.md.sec 			<- lapply(tab_names[as.numeric(outcomes_qes.md.sec$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C45:J64"))
secondary_outcomes.md 		<- data.frame(data.table::rbindlist(tibble_qes.md.sec, idcol='ID')) # merges everything to one dataframe

# Merge results for mean-difference data (md)
tibble_qes.md 				<- list(tibble_qes.md.prim, tibble_qes.md.sec) 		# facilitates data handling in a loop later
outcomes_all.md 			<- list(primary_outcomes.md, secondary_outcomes.md) # facilitates data handling in a loop later ## NOT USED, REMOVE?
outcomes_qes.md 			<- list(outcomes_qes.md.prim, outcomes_qes.md.sec)


# ==================================================================================================
# A. Extract SMD of RCTs with baseline and follow-up data via loop
# General settings
category_condition 	<- c("baseline", "follow-up")
studies2exclude 	<- c(outcomes_rct.md.prim$author, outcomes_rct.md.sec$author, # these are the studies that are problematic for some reason
					"Heinonen et al. 1989", "King et al. 2009", "Nutt et al. 2007","Sivertsen et al. 1989", 
					"Macht et al. 2000", "Kadkhodaie et al. 2019", "Bara-Jimenez et al. 2003") # list of studies that should be likewise reported 		

# Formula used 
df_transpose <- function(df) { # function intended to transpose a matrix
  df %>% 
    tidyr::pivot_longer(-1) %>%
    tidyr::pivot_wider(names_from = 1, values_from = value)
}

# Pre-allocate dataframes to fill later:
df_meta_TRT.blfu			<- data.frame(matrix(, nrow = 0, ncol = 11)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT.blfu) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")
df_meta_CTRL.blfu			<- df_meta_TRT.blfu
df_doublecheck 				<- df_meta_TRT.blfu 									# list to be filled while looping over results

# Start extracting data:
writeLines(sprintf("%s", strrep("=",40)))
writeLines("\nExtracting results from pre-post controlled RCT studies (primary AND secondary outcomes):")

for (o in 1:2){ # loop through primary and secondary outcomes
	outcomes_to_test 	<- data.frame(outcomes_rct.blfu[o])
	idx_studies2exclude <- which(outcomes_rct.blfu[[o]]$author %in% studies2exclude, arr.ind=TRUE)

	for (i in setdiff( 1:dim(outcomes_to_test)[1], idx_studies2exclude)){ # loop through all available studies and fill df_meta
		writeLines(sprintf("Processing study: %s", outcomes_to_test$author[i]))

		tibble_temp = tibble_rct.blfu[[o]]										# dataframe of either primary or secondary outcomes
		idx_study_total 	<- as.numeric(outcomes_to_test$type[i]) 		# index of patient in the entire list (cf. {first_authors} above)
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 11))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")

		results_temp 		<- tibble_temp[[i]] 							# the results of the outcomes
		outcome_temp		<- outcomes_to_test %>% 						# 
			select(matches("outcome")) %>% 
			slice(i) %>% t()
		colnames(outcome_temp) <- "outcome"

		data_of_interest 	<- data.frame(outcome_temp) %>% 		# all outcomes should be sorted as *outcome*__*comparator*_time, e.g. 'UPDRS_item20_21__pramipexole_baseline'
			filter_all(any_vars(!is.na(.))) %>% 							# the next few lines are intended to get some details to work with
			separate(outcome, into=c("name_outcome", "comparator_time"), sep="__", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator_time) %>%
			separate(comparator_time, into=c("comparator", "time"), sep="_", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator, time) %>%
			filter(across(name_outcome, ~ !grepl('adverse', .)))
		available_interventions <- unique(data_of_interest$comparator)  	# lists the available interventions to loop through
		
		df_temp$study = i													# Fill temporary dataframe {df_temp} (cf. line 271) with content
		df_temp$source <- outcomes_to_test$author[i]
		
		if (length(available_interventions) == 1) {
			cat(sprintf("\t\t... study: %s seems to be lacking pre-post design. Adding to manual list", outcomes_to_test$author[i]))
			df_doublecheck <- rbind(df_doublecheck, df_temp)
			next
		} else if (length(available_interventions)>1) {
		
			# loop through  possible combinations of interventions (cf. combn) and append available data
			all_effects 			<- combn(available_interventions, 2) # list of possible effects resulting from available interventions (cf. available_interventions)
			extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																	sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
			extracted_qualsyst 		<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																	sheet = x, col_names=c("QS1", "QS2", "QS3", "QS4"), range = "C69:F69")) # extracts QualSyst values
			extracted_qualsyst 		<- extracted_qualsyst %>% bind_rows(., .id="sheet")
			df_temp$qualsyst		<- extracted_qualsyst$QS1 / extracted_qualsyst$QS2 
			df_numbers 				<- df_transpose(extracted_demographics[[1]])
			colnames(df_numbers)	<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
			df_temp_multiple 		<- df_temp 	# necessary to maintain the structure for all combinations
			
			for (k in 1:dim(all_effects)[2]){ 	# loop through all possible combinations, cf. all_effects <- combn(columns_of_interest$outcome, 2)
				
				for (c in 1:2){ # Loop through pairs of comparators (treatment vs. control)
					# TREATMENT (-> df_meta_TRT, c == 1) or CONTROL (-> df_meta_CTRL, c == 2)
					df_temp_multiple 			<- df_temp 	# necessary to maintain the structure for all combinations in case there are more than 1 comparisons
					df_temp_multiple$type 		<- 'RCT'
					columns_of_interest 		<- data.frame(outcome_temp) %>% filter(rowSums(across(everything(), ~grepl(all_effects[,k][1], .x))) > 0)
					
					# Fill dataframe with content extracted from the results of primary outcomes and from demographics
					df_temp_multiple$comparator <- all_effects[,k][c] # str_split(columns_of_interest$outcome[idx_comparator1], '_')[[1]][1]
					
					idx_numbers 				<- which(df_numbers$intervention==df_temp_multiple$comparator) 					# column in which data is saved
					if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 	# when there are no 'arms' all data should be listed as 'total'
					temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator, '_', category_condition[1]), .x))) > 0)
					df_temp_multiple$m_pre 		<- as.numeric(temp_vals$mean)
					df_temp_multiple$sd_pre 	<- as.numeric(temp_vals$sd)
					df_temp_multiple$ni 		<- as.numeric(df_numbers$included[idx_numbers])
					
					# FOLLOW-UP (-> df_meta_post)
					columns_of_interest 		<- data.frame(outcome_temp) %>% filter(rowSums(across(everything(), ~grepl(category_condition[2], .x))) > 0)

					df_temp_multiple$comparator <- all_effects[,k][c] #str_split(columns_of_interest$outcome[idx_comparator2], '_')[[1]][1]
					idx_numbers 				<- which(df_numbers$intervention==df_temp_multiple$comparator) # column in which data is saved
					if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 
					temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator, '_', category_condition[2]), .x))) > 0)
					df_temp_multiple$sd_post	<- as.numeric(temp_vals$sd)
					df_temp_multiple$m_post 	<- as.numeric(temp_vals$mean)
					df_temp_multiple$ni 		<- as.numeric(df_numbers$included[idx_numbers])
					
					if (c == 1) {
						df_meta_TRT.blfu 		<- rbind(df_meta_TRT.blfu, df_temp_multiple) # append to list	
					} else {
						df_meta_CTRL.blfu 		<- rbind(df_meta_CTRL.blfu, df_temp_multiple)
					}
				}
			}
		} else {
			writeLines(sprintf("\t\tNot sure what the problem is with study: %s", outcomes_to_test$author[i]))
			df_doublecheck <- rbind(df_doublecheck, df_temp)

		}
	}
}
cat("\n...Done!")
writeLines(sprintf("\n%s", strrep("=",40)))

# Clean up the results and estimate "SMCR", in order to standardize all resultsso that it can be used later:
df_meta_CTRL.blfu 		<- df_meta_CTRL.blfu %>% drop_na(m_pre, m_post, sd_pre, sd_post)
df_meta_CTRL.blfu$ri 	<- 0 # correlation matrix set to 0
df_meta_TRT.blfu 		<- df_meta_TRT.blfu %>% drop_na(m_pre, m_post, sd_pre, sd_post)
df_meta_TRT.blfu$ri 	<- 0 # correlation matrix set to 0

only_placebocontrolled = TRUE
if (only_placebocontrolled){
	idx_remove 			<- c()
	idx_placebo 		<- which(df_meta_TRT.blfu$comparator=="placebo")								# removes placebo in the TRT group
	comparators 		<- data.frame(cbind(df_meta_TRT.blfu$comparator, df_meta_CTRL.blfu$comparator))
	idx_nocontrol 		<- comparators %>% rowid_to_column() %>% 
							filter(!if_any(everything(),~str_detect(., c("placebo|control")))) %>% 
							select(rowid)															# selects where there is no placebo/control group
	idx_remove 			<- c(idx_remove, idx_placebo, idx_nocontrol$rowid)
	df_meta_TRT.blfu 	<- df_meta_TRT.blfu[setdiff(1:dim(df_meta_TRT.blfu)[1], idx_remove),]
	df_meta_CTRL.blfu 	<- df_meta_CTRL.blfu[setdiff(1:dim(df_meta_CTRL.blfu)[1], idx_remove),]
}

datC.rct.blfu 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_meta_CTRL.blfu)
datT.rct.blfu 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_meta_TRT.blfu)
dat.rct.blfu 			<- data.frame(yi = datT.rct.blfu$yi - datC.rct.blfu$yi, vi = datT.rct.blfu$vi + datC.rct.blfu$vi, slab=datC.rct.blfu$source)
dat.rct.blfu$treatment	<- df_meta_TRT.blfu$comparator
dat.rct.blfu$study_type	<- df_meta_TRT.blfu$type
dat.rct.blfu$qualsyst	<- df_meta_TRT.blfu$qualsyst
dat.rct.blfu$ni			<- df_meta_TRT.blfu$ni


# Some studies must be added manually:
# a. Rascol et al. 1988 (only values pre- and post available so that the former are the baseline)
df_temp <- df_doublecheck
df_temp$comparator <- "selegiline"
df_temp$m_pre 	<- 1.7
df_temp$sd_pre 	<- 0.4
df_temp$m_post 	<- 0.9
df_temp$sd_post <- 0.2
df_temp$ni 		<- 16
df_temp$qualsyst<- 16/20
df_temp$study_type <- "RCT"
dat_manual 		<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=1, slab=source, data=df_temp)
df_temp  			<- dat.rct.blfu[1,]
df_temp$slab  		<- dat_manual$source
df_temp$yi 			<- dat_manual$yi
df_temp$vi 			<- dat_manual$vi
df_temp$treatment 	<- dat_manual$comparator
df_temp$study_type  <- dat_manual$study_type
df_temp$qualsyst 	<- dat_manual$qualsyst
df_temp$ni 			<- 15

dat.rct.blfu <- rbind(dat.rct.blfu, df_temp)
			
# ==================================================================================================
# B. Extract data from RCT with mean-differences

studies2exclude 	<- c()

# Pre-allocate dataframes to fill later:
df_meta_TRT.md 			<- data.frame(matrix(, nrow = 0, ncol = 11)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT.md)<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")
df_meta_CTRL.md			<- df_meta_TRT.md
df_doublecheck.md 		<- df_meta_TRT.md 								# list to be filled while looping over results

writeLines(sprintf("%s", strrep("=",40)))
writeLines("\nExtracting results from RCT (primary AND secondary outcomes) with disclosed mean-differences:")

for (o in 1:2){ # loop through primary and secondary outcomes

	outcomes_to_test 	<- data.frame(outcomes_rct.md[[o]])
	idx_studies2exclude <- which(outcomes_rct.md[[o]]$author %in% studies2exclude, arr.ind=TRUE)
	for (i in setdiff( 1:dim(outcomes_to_test)[1], idx_studies2exclude)){ # loop through all studies encountered before
		writeLines(sprintf("Processing study: %s - number %s", outcomes_to_test$author[i], i))
		
		tibble_temp = tibble_rct.md[[o]]										# dataframe of either primary or secondary outcomes
		idx_study_total 	<- as.numeric(outcomes_to_test$type[i]) 		# index of patient in the entire list (cf. {first_authors} above)
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 11))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")

		results_temp 		<- tibble_temp[[i]] 							# the results of the outcomes
		outcome_temp		<- outcomes_to_test %>% 						# 
			select(matches("outcome")) %>% 
			slice(i) %>% t()
		colnames(outcome_temp) <- "outcome"

		data_of_interest 	<- data.frame(outcome_temp) %>% 		# all outcomes should be sorted as *outcome*__*comparator*_time, e.g. 'UPDRS_item20_21__pramipexole_baseline'
			filter_all(any_vars(!is.na(.))) %>% 							# the next few lines are intended to get some details to work with
			separate(outcome, into=c("name_outcome", "comparator_time"), sep="__", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator_time) %>%
			separate(comparator_time, into=c("comparator", "arm"), sep="_", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator, arm) %>%
			filter(across(name_outcome, ~ !grepl('adverse', .)))
		available_interventions <- unique(data_of_interest$comparator)  	# lists the available interventions to loop through
			
		df_temp$study = i													# Fill temporary dataframe {df_temp} (cf. line 271) with content
		df_temp$source <- outcomes_to_test$author[i]

		if (length(available_interventions) != 2) {
				cat(sprintf("\t\t... study: %s seems to be lacking a control group. Adding to manual list", outcomes_to_test$author[i]))
				df_doublecheck.md <- rbind(df_doublecheck.md, df_temp)
				next
		} else { # if (length(available_interventions)==2) # two arms available 
			
				# loop through  possible combinations of interventions (cf. combn) and append available data
				all_effects 			<- combn(available_interventions, 2) # list of possible effects resulting from available interventions (cf. available_interventions)
				extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																		sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
				extracted_qualsyst 		<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																	sheet = x, col_names=c("QS1", "QS2", "QS3", "QS4"), range = "C69:F69")) # extracts QualSyst values
				extracted_qualsyst 		<- extracted_qualsyst %>% bind_rows(., .id="sheet")
				df_temp$qualsyst		<- extracted_qualsyst$QS1 / extracted_qualsyst$QS2 

				df_numbers 				<- df_transpose(extracted_demographics[[1]])
				colnames(df_numbers)	<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
				df_temp_multiple 		<- df_temp 	# necessary to maintain the structure for all combinations
				
				for (k in 1:dim(all_effects)[1]){ 	# loop through all possible comparators
						# all data is sorted so that TREATMENT is the first and comparator the second line
						df_temp_multiple 			<- df_temp 	# necessary to maintain the structure for all combinations in case there are more than 1 comparisons
						df_temp_multiple$type 		<- 'RCT'
						
						columns_of_interest 		<- data.frame(outcome_temp) %>% filter(rowSums(across(everything(), ~grepl(all_effects[k], .x))) > 0)
						
						# Fill dataframe with content extracted from the results of primary outcomes and from demographics
						df_temp_multiple$comparator <- all_effects[k] # str_split(columns_of_interest$outcome[idx_comparator1], '_')[[1]][1]
						
						idx_numbers 				<- which(df_numbers$intervention==df_temp_multiple$comparator) 					# column in which data is saved
						if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 	# when there are no 'arms' all data should be listed as 'total'
						#temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator, '_', "mean-difference"), .x))) > 0)
						temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(sprintf("%s|%s", paste0(df_temp_multiple$comparator, '_', "mean-difference"), paste0(df_temp_multiple$comparator, '_', "adjusted-mean-difference")), .x))) > 0)
						# filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .))) %>%

						# filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .))) %>%
						df_temp_multiple$m_pre 		<- as.numeric(temp_vals$mean)
						df_temp_multiple$m_post 	<- 0
						df_temp_multiple$sd_pre 	<- as.numeric(temp_vals$baseline_sd)
						if (is.na(df_temp_multiple$sd_pre)) {
							df_temp_multiple$sd_pre <- temp_vals$sd}
						df_temp_multiple$sd_pre 	<- as.numeric(temp_vals$sd)
						df_temp_multiple$sd_post 	<- 0
						df_temp_multiple$ni 		<- as.numeric(df_numbers$included[idx_numbers])
											
						if (k == 1) {
							df_meta_TRT.md 			<- rbind(df_meta_TRT.md, df_temp_multiple) # append to list	
						} else {
							df_meta_CTRL.md			<- rbind(df_meta_CTRL.md, df_temp_multiple)
						}
				}
			}	
	}
}

# Estimate "SMCR" which can be used later:
df_meta_CTRL.md <- df_meta_CTRL.md %>% drop_na(m_pre, m_post)
df_meta_CTRL.md$m_pre <- df_meta_CTRL.md$m_pre * (-1)
df_meta_CTRL.md$ri 	<- 0 # correlation matrix set to 0

df_meta_TRT.md <- df_meta_TRT.md %>% drop_na(m_pre, m_post)
df_meta_TRT.md$m_pre <- df_meta_TRT.md$m_pre * (-1)
df_meta_TRT.md$ri <- 0

only_placebocontrolled = TRUE
if (only_placebocontrolled){
	idx_remove 			<- c()
	idx_placebo 		<- which(df_meta_TRT.md$comparator=="placebo")								# removes placebo in the TRT group
	comparators 		<- data.frame(cbind(df_meta_TRT.md$comparator, df_meta_CTRL.md$comparator))
	idx_nocontrol 		<- comparators %>% rowid_to_column() %>% 
							filter(!if_any(everything(),~str_detect(., c("placebo|control")))) %>% 
							select(rowid)															# selects where there is no placebo/control group
	idx_remove 			<- c(idx_remove, idx_placebo, idx_nocontrol$rowid)
	df_meta_TRT.md 	<- df_meta_TRT.md[setdiff(1:dim(df_meta_TRT.md)[1], idx_remove),]
	df_meta_CTRL.md 	<- df_meta_CTRL.md[setdiff(1:dim(df_meta_CTRL.md)[1], idx_remove),]
}

datC.rct.md 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_meta_CTRL.md)
datT.rct.md 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_meta_TRT.md)
dat.rct.md 				<- data.frame(yi = datT.rct.md$yi - datC.rct.md$yi, vi = datT.rct.md$vi + datC.rct.md$vi, slab=datC.rct.md$source)
dat.rct.md$treatment	<- df_meta_TRT.md$comparator
dat.rct.md$study_type	<- df_meta_TRT.md$type
dat.rct.md$qualsyst		<- df_meta_TRT.md$qualsyst
dat.rct.md$ni			<- df_meta_TRT.md$ni


# Some studies must be added manually:
# b. Ziegler et al. 2003 (only p-values resulting from t-test were available)
df_temp 		<- df_meta_TRT.md  %>% rowid_to_column() %>% filter(if_any(everything(),~str_detect(., c("Ziegler"))))
df_temp$pvalue 	<- .088
df_temp$tval	<- NA
df_temp$dval 	<- NA
df_temp$tval 	<- replmiss(df_temp$tval, with(df_temp, -1 * qt(pvalue/2, df=ni-1, lower.tail=FALSE)))
df_temp$dval 	<- replmiss(df_temp$dval, with(df_temp, tval * sqrt(1/ni + 1/ni)))
df_temp$yi 		<- NA
df_temp$yi 		<- replmiss(df_temp$yi, with(df_temp, (1 - 3/(4*(ni-1) - 1)) * dval))
df_temp$vi 		<- NA
df_temp$vi 		<- replmiss(df_temp$vi, with(df_temp, 1/ni+ 1/ni + yi^2/(2*(ni+ni))))
dat.rct.md$yi[df_temp$rowid] <- df_temp$yi
dat.rct.md$vi[df_temp$rowid] <- df_temp$vi
dat.rct.md$ni[df_temp$rowid] <- df_temp$ni


# c. Malsch et al. 2001
df_temp 			<- dat.rct.md[1:2,]
df_temp$slab 		<- c("Malsch et al. 2001", "Malsch et al. 2001.1")
df_temp$treatment 	<- c("amantadine", "budipine")
df_temp$study_type 	<- c("RCT", "RCT")
df_temp$qualsyst 	<- c(20/22, 20/22)
df_temp$ni			<- c(26, 27)
df_temp 			<- escalc(measure="SMCR", m1i=c(-.58, -1.06), m2i=c(0,0), sd1i=c(1.4, 1.8), ni=c(26, 27), ri=c(0,0), slab=slab, data=df_temp)

dat.rct.md <- rbind(dat.rct.md, df_temp)

# c. Nomoto et al. 2018 (only p-values resulting from t-test were available)
df_temp 		<- df_meta_TRT.md  %>% rowid_to_column() %>% filter(if_any(everything(),~str_detect(., c("Nomoto"))))
df_temp$pvalue 	<- c(.0005, .0001)
df_temp$tval	<- c(NA, NA)
df_temp$dval 	<- c(NA, NA)
df_temp$tval 	<- replmiss(df_temp$tval, with(df_temp, -1 * qt(pvalue/2, df=ni-1, lower.tail=FALSE)))
df_temp$dval 	<- replmiss(df_temp$dval, with(df_temp, tval * sqrt(1/ni + 1/ni)))
df_temp$yi 		<- c(NA, NA)
df_temp$yi 		<- replmiss(df_temp$yi, with(df_temp, (1 - 3/(4*(ni-1) - 1)) * dval))
df_temp$vi 		<- c(NA, NA)
df_temp$vi 		<- replmiss(df_temp$vi, with(df_temp, 1/ni+ 1/ni + yi^2/(2*(ni+ni))))
dat.rct.md$yi[df_temp$rowid] <- df_temp$yi
dat.rct.md$vi[df_temp$rowid] <- df_temp$vi
dat.rct.md$ni[df_temp$rowid] <- df_temp$ni

dat_results <- rbind(dat.rct.blfu, dat.rct.md)

# Add manually: Friedman et al. 1997 (ask Goetz for baseline data)


# ==================================================================================================
# C. Extract SMD of QES with baseline and follow-up data via loop
# General settings

studies2exclude 	<- c(outcomes_qes.md.prim$author, outcomes_qes.md.sec$author,
							"Samotus et al. 2020", "Spieker et al. 1995", "Barbagallo et al. 2018")

# Pre-allocate dataframes to fill later:
df_meta_TRT.qes 		<- data.frame(matrix(, nrow = 0, ncol = 11)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT.qes) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")
df_doublecheck 			<- df_meta_TRT.qes 									# list to be filled while looping over results

# Start extracting data:
writeLines(sprintf("%s", strrep("=",40)))
writeLines("\nExtracting results from pre-post QES studies (primary AND secondary outcomes):")

for (o in 1:2){ # loop through primary and secondary outcomes
	outcomes_to_test 	<- data.frame(outcomes_qes.blfu[o])
	idx_studies2exclude <- which(outcomes_qes.blfu[[o]]$author %in% studies2exclude, arr.ind=TRUE)

	for (i in setdiff( 1:dim(outcomes_to_test)[1], idx_studies2exclude)){ # loop through all available studies and fill df_meta
		writeLines(sprintf("Processing study: %s", outcomes_to_test$author[i]))

		tibble_temp = tibble_qes.blfu[[o]]										# dataframe of either primary or secondary outcomes
		idx_study_total 	<- as.numeric(outcomes_to_test$type[i]) 		# index of patient in the entire list (cf. {first_authors} above)
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 11))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type", "qualsyst")

		results_temp 		<- tibble_temp[[i]] 							# the results of the outcomes
		outcome_temp		<- outcomes_to_test %>% 						# 
			select(matches("outcome")) %>% 
			slice(i) %>% t()
		colnames(outcome_temp) <- "outcome"

		data_of_interest 	<- data.frame(outcome_temp) %>% 		# all outcomes should be sorted as *outcome*__*comparator*_time, e.g. 'UPDRS_item20_21__pramipexole_baseline'
			filter_all(any_vars(!is.na(.))) %>% 							# the next few lines are intended to get some details to work with
			separate(outcome, into=c("name_outcome", "comparator_time"), sep="__", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator_time) %>%
			separate(comparator_time, into=c("comparator", "time"), sep="_", fill = "right", remove = FALSE) %>%
			select(name_outcome, comparator, time) %>%
			filter(across(name_outcome, ~ !grepl('adverse', .)))
		available_interventions <- unique(data_of_interest$comparator)  	# lists the available interventions to loop through
		
		df_temp$study = i													# Fill temporary dataframe {df_temp} (cf. line 271) with content
		df_temp$source <- outcomes_to_test$author[i]
		extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
												sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
		extracted_qualsyst 		<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																	sheet = x, col_names=c("QS1", "QS2", "QS3", "QS4"), range = "C69:F69")) # extracts QualSyst values
		extracted_qualsyst 		<- extracted_qualsyst %>% bind_rows(., .id="sheet")
		df_temp$qualsyst		<- extracted_qualsyst$QS1 / extracted_qualsyst$QS2 

		
		df_numbers 				<- df_transpose(extracted_demographics[[1]])
		colnames(df_numbers)	<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
		
		if (length(available_interventions) == 1) {
			
			df_temp$type 		<- 'QES'
			columns_of_interest <- data.frame(outcome_temp) %>% filter(rowSums(across(everything(), ~grepl(available_interventions[1], .x))) > 0)
			df_temp$comparator 	<- available_interventions
			idx_numbers 		<- which(df_numbers$intervention==df_temp$comparator) 					# column in which data is saved
			if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 	# when there are no 'arms' all data should be listed as 'total'
		
			# BASELINE
			temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp$comparator, '_', category_condition[1]), .x))) > 0)
			df_temp$m_pre 		<- as.numeric(temp_vals$mean)
			df_temp$sd_pre 		<- as.numeric(temp_vals$sd)
			df_temp$ni 			<- as.numeric(df_numbers$included[idx_numbers])

			# FOLLOW-UP
			temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp$comparator, '_', category_condition[2]), .x))) > 0)
			df_temp$m_post 		<- as.numeric(temp_vals$mean)
			df_temp$sd_post 	<- as.numeric(temp_vals$sd)
			df_temp$ni 			<- as.numeric(df_numbers$included[idx_numbers])			
			
			df_meta_TRT.qes 	<- rbind(df_meta_TRT.qes, df_temp)
		} else if (length(available_interventions)>1) {

			# loop through  possible combinations of interventions (cf. combn) and append available data
			all_effects 			<- combn(available_interventions, 2) # list of possible effects resulting from available interventions (cf. available_interventions)
			#extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
			#														sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
			#df_numbers 				<- df_transpose(extracted_demographics[[1]])
			colnames(df_numbers)	<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
			df_temp_multiple 		<- df_temp 	# necessary to maintain the structure for all combinations
			
			for (k in 1:dim(all_effects)[1]){ 	# loop through all possible combinations, cf. all_effects <- combn(columns_of_interest$outcome, 2)
				
				df_temp_multiple$type 		<- 'QES'
				columns_of_interest 		<- data.frame(outcome_temp) %>% filter(rowSums(across(everything(), ~grepl(available_interventions[1], .x))) > 0)
				df_temp_multiple$comparator <- all_effects[,1][k]
				idx_numbers 				<- which(df_numbers$intervention==df_temp_multiple$comparator) 					# column in which data is saved
				if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 	# when there are no 'arms' all data should be listed as 'total'
			
				# BASELINE
				temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator, '_', category_condition[1]), .x))) > 0)
				df_temp_multiple$m_pre 		<- as.numeric(temp_vals$mean)
				df_temp_multiple$sd_pre 	<- as.numeric(temp_vals$sd)
				df_temp_multiple$ni 		<- as.numeric(df_numbers$included[idx_numbers])

				# FOLLOW-UP
				temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator, '_', category_condition[2]), .x))) > 0)
				df_temp_multiple$m_post 	<- as.numeric(temp_vals$mean)
				df_temp_multiple$sd_post 	<- as.numeric(temp_vals$sd)
				df_temp_multiple$ni 		<- as.numeric(df_numbers$included[idx_numbers])			
				
				df_meta_TRT.qes 			<- rbind(df_meta_TRT.qes, df_temp_multiple)
			}
		} else {
			writeLines(sprintf("\t\tNot sure what the problem is with study: %s", outcomes_to_test$author[i]))
			df_doublecheck <- rbind(df_doublecheck, df_temp)

		}
	}
}
cat("\n...Done!")
writeLines(sprintf("\n%s", strrep("=",40)))

# Estimate "SMCR" which can be used later:
df_meta_TRT.qes <- df_meta_TRT.qes %>% drop_na(m_pre, m_post, sd_pre)
df_meta_TRT.qes$ri 	<- 0 # correlation matrix set to 0

datT.qes 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_meta_TRT.qes)
dat.qes 			<- data.frame(yi = datT.qes$yi, vi = datT.qes$vi, slab=datT.qes$source)
dat.qes$treatment	<- df_meta_TRT.qes$comparator
dat.qes$study_type	<- df_meta_TRT.qes$type
dat.qes$qualsyst	<- df_meta_TRT.qes$qualsyst
dat.qes$ni			<- df_meta_TRT.qes$ni

# e. Barbagallo et al. 2018
df_tempT 			<- datT.rct.md[1,] 	
df_tempT$comparator <- "apomorphine"	
df_tempT$source 	<- "Barbagallo et al. 2018"
df_tempT$m_pre 		<- 4.9
df_tempT$m_post 	<- 2.1
df_tempT$sd_pre 	<- 1.8
df_tempT$sd_post 	<- 1.0
df_tempT$ni			<- 15
df_tempC			<- df_tempT	
df_tempC$comparator <- "placebo"	
df_tempC$m_pre 		<- 4.9
df_tempC$m_post 	<- 3.3
df_tempC$sd_pre 	<- 1.8
df_tempC$sd_post 	<- 1.8
df_tempC$ni			<- 15

df_tempT 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=0, slab=source, data=df_tempT)
df_tempC 			<- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=0, slab=source, data=df_tempC)
df_temp 				<- data.frame(yi = df_tempT$yi - df_tempC$yi, vi = df_tempT$vi + df_tempC$vi, slab=df_tempC$source)

df_temp$treatment	<- df_tempT$comparator
df_temp$study_type	<- df_tempT$type
df_temp$qualsyst	<- 15/28
df_temp$ni			<- df_tempT$ni

dat.qes <- rbind(dat.qes, df_temp)

# f. Altman et al. 2012
df_temp 			<- dat.qes[1,]
df_temp$slab 		<- "Altman et al. 2012"
df_temp$treatment 	<- "caffeine"
df_temp$study_type 	<- "QES"
df_temp$qualsyst 	<- 16/20
df_temp$ni			<- 48
df_temp 			<- escalc(measure="SMCR", m1i=-1.18, m2i=0, sd1i=3.2, ni=48, ri=0, slab=slab, data=df_temp)

dat.qes <- rbind(dat.qes, df_temp)

# g. Samotus et al. 2020
df_temp 			<- dat.qes[1,]
df_temp$slab 		<- "Samotus et al. 2020"
df_temp$treatment 	<- "incobotulinumtoxin"
df_temp$study_type 	<- "QES"
df_temp$qualsyst 	<- 11/22
df_temp$ni			<- 48
df_temp 			<- escalc(measure="SMCR", m1i=-1,55, m2i=0, sd1i=1.34, ni=48, ri=0, slab=slab, data=df_temp)

dat.qes <- rbind(dat.qes, df_temp)

dat.qes 			<- dat.qes %>% drop_na()
dat_results <- rbind(dat_results, dat.qes)

dat_results %>% write.csv(file.path(wdir, "results", "results_meta-analysis.csv"), row.names = F) # save results to file
