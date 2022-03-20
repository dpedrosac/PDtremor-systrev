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
tibble_rct.md 				<- list(tibble_qes.md.prim, tibble_qes.md.sec) 		# facilitates data handling in a loop later
outcomes_all.md 			<- list(primary_outcomes.md, secondary_outcomes.md) # facilitates data handling in a loop later ## NOT USED, REMOVE?
outcomes_qes.md 			<- list(outcomes_qes.md.prim, outcomes_qes.md.sec)


# ==================================================================================================
# A. Extract SMD of RCTs with baseline and follow-up data via loop
# General settings
category_condition 	<- c("baseline", "follow-up")
categories 			<- c("levodopa", "dopamine agonists", "mao-inhibitors", "amantadin", "physical exercise", "other") 	# categories to be used as subgroups

studies2exclude 	<- c(outcomes_rct.md.prim$author, outcomes_rct.md.sec$author, "Bara-Jimenez et al. 2003", # list of studies that need manual checks 																				# these are the studies that are problematic for some reason
					"Bergamasco et al. 2000", "Eberhardt et al. 1990", "Heinonen et al. 1989", "King et al. 2009", "Nomoto et al. 2018", "Olanow et al. 1987", 
					"Tortolero et al. 2004", "Kadkhodaie et al. 2019", "Macht et al. 2000", # "Navan et al. 2003.1", "Navan et al. 2003",
					"Nutt et al. 2007", #"Olson et al. 1997",
					"Sivertsen et al. 1989") #, "Spieker et al. 1999")

# Formula used in the next loop 
df_transpose <- function(df) { # function intended to transpose a matrix
  df %>% 
    tidyr::pivot_longer(-1) %>%
    tidyr::pivot_wider(names_from = 1, values_from = value)
}

# Pre-allocate dataframes to fill later:
df_meta_TRT 			<- data.frame(matrix(, nrow = 0, ncol = 10)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")
df_meta_CTRL			<- df_meta_TRT
df_doublecheck 			<- df_meta_TRT 									# list to be filled while looping over results

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
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 10))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")

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
						df_meta_TRT 			<- rbind(df_meta_TRT, df_temp_multiple) # append to list	
					} else {
						df_meta_CTRL 			<- rbind(df_meta_CTRL, df_temp_multiple)
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

# ==================================================================================================
# B. Extract data from RCT with mean-differences

#TODO uincorporate that baseline sd is used to standardize!

studies2exclude 	<- c("Postuma et al. 2012")

# Formula used in the next loop 
df_transpose <- function(df) { # function intended to transpose a matrix
  df %>% 
    tidyr::pivot_longer(-1) %>%
    tidyr::pivot_wider(names_from = 1, values_from = value)
}

# Pre-allocate dataframes to fill later:
df_meta_TRT.md 			<- data.frame(matrix(, nrow = 0, ncol = 10)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT.md)<- c("study", "type", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")
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
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 10))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")

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
				df_doublecheck <- rbind(df_doublecheck, df_temp)
				next
		} else { # if (length(available_interventions)==2) # two arms available 
			
				# loop through  possible combinations of interventions (cf. combn) and append available data
				all_effects 			<- combn(available_interventions, 2) # list of possible effects resulting from available interventions (cf. available_interventions)
				extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
																		sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
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

# Add manually: Ziegler et al. 2003 (p-values available with t-test), Friedman et al. 1997 (ask Goetz for baseline data), Jankovic et al. 2014 (statistical help with ANCOVA),
# Malsch et al. 2001 (baseline_sd available), Nomoto et al. 2018 (p-values available with t-test)

# ==================================================================================================
# C. Extract SMD of QES with baseline and follow-up data via loop
# General settings

studies2exclude 	<- c(outcomes_qes.md.prim$author, outcomes_qes.md.sec$author, "Milanov (2001)",
							"Samotus et al. 2020", "Spieker et al. 1995") #, "Bara-Jimenez et al. 2003", # list of studies that need manual checks 																				# these are the studies that are problematic for some reason

# Formula used in the next loop 
df_transpose <- function(df) { # function intended to transpose a matrix
  df %>% 
    tidyr::pivot_longer(-1) %>%
    tidyr::pivot_wider(names_from = 1, values_from = value)
}

# Pre-allocate dataframes to fill later:
df_meta_TRT.qes 		<- data.frame(matrix(, nrow = 0, ncol = 10)) 	# dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta_TRT.qes) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")
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
		df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 10))	# temporary dataframe
		colnames(df_temp) 	<- c("study", "source", "comparator", "m_pre", "m_post", "sd_pre", "sd_post", "ni", "ri", "type")

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
			sprintf("Not yet defined, please use i = %s fo debugging", i)
			next 
			# loop through  possible combinations of interventions (cf. combn) and append available data
			all_effects 			<- combn(available_interventions, 2) # list of possible effects resulting from available interventions (cf. available_interventions)
			#extracted_demographics 	<- lapply(tab_names[as.numeric(outcomes_to_test$type[i])], function(x) read_excel(results_file, 
			#														sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31")) # extracts demographics for all interventions
			#df_numbers 				<- df_transpose(extracted_demographics[[1]])
			colnames(df_numbers)	<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
			df_temp_multiple 		<- df_temp 	# necessary to maintain the structure for all combinations
			
			for (k in 1:dim(all_effects)[2]){ 	# loop through all possible combinations, cf. all_effects <- combn(columns_of_interest$outcome, 2)
				
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
				temp_vals 					<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp$comparator, '_', category_condition[2]), .x))) > 0)
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



# ==================================================================================================
# ==================================================================================================
# ==================================================================================================
# ==================================================================================================
# ==================================================================================================
# Start withe meta-analysis for RCTs

# This is the format the data frame should be put into, to further work with the data in the metafor
# package (cf.: https://www.metafor-project.org/doku.php/tips:assembling_data_smd)

#  study             source n1i m1i sd1i n2i m2i sd2i
#1     1          Edinburgh 155  55   47 156  75   64
#2     2     Orpington-Mild  31  27    7  32  29    4
#3     3 Orpington-Moderate  75  64   17  71 119   29
#4     4   Orpington-Severe  18  66   20  18 137   48
#5     5      Montreal-Home   8  14    8  13  18   11
#6     6  Montreal-Transfer  57  19    7  52  18    4
#7     7          Newcastle  34  52   45  33  41   34
#8     8               Umea 110  21   16 183  31   27
#9     9            Uppsala  60  30   27  52  23   20

# First extract all data from the before after results (baseline vs. follow_up)
outcomes_rct <- outcomes_wideRCT %>% filter(!is.na(primary_outcome1)) %>% 
	rownames_to_column("type") %>% 
	filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .))) %>%
	head()
	
idx_outcomes_rct1 <- outcomes_wide %>% # idx_outcomes_rct1 renamed to outcomes_rct.blfu.prim
	rownames_to_column("type") %>%
	filter(study_type=="Randomised-controlled trial (RCT)") %>%
	filter(across(primary_outcome1, ~ grepl('baseline|follow-up', .)))%>% 
	separate(primary_outcome1, "__",
                into = c("the_rest", "outcome1"), 
                remove = TRUE) %>% 
	separate(primary_outcome2, "__",
                into = c("the_rest", "outcome2"), 
                remove = TRUE) %>%
	separate(primary_outcome3, "__",
                into = c("the_rest", "outcome3"), 
                remove = TRUE) %>%
	separate(primary_outcome4, "__",
                into = c("the_rest", "outcome4"), 
                remove = TRUE) %>%
	separate(primary_outcome5, "__",
                into = c("the_rest", "outcome5"), 
                remove = TRUE) %>%
	separate(primary_outcome6, "__",
                into = c("the_rest", "outcome6"), 
                remove = TRUE) %>%
	separate(primary_outcome7, "__",
                into = c("the_rest", "outcome7"), 
                remove = TRUE) %>%
	separate(primary_outcome8, "__",
                into = c("the_rest", "outcome8"), 
                remove = TRUE) %>% select(-any_of("the_rest"))

extracted_results1_p 	<- lapply(tab_names[as.numeric(idx_outcomes_rct1$type)], function(x) read_excel(results_file, 
															sheet = x, col_names=c('outcome', 'mean', 'sd', 'sem', 'ci_low', 'ci_high', 'pvalue', 'baseline_sd'), range = "C35:H44"))

results_first_attempt 			<- data.frame(data.table::rbindlist(extracted_results1_p, idcol='ID')) # merges everything to one dataframe


# ==================================================================================================
# Extract SMD via loop; this is a somehow very manual loop, which in turn offers some flexibility when studies are very individual

df_transpose <- function(df) { # function intended to transpose a matrix
  df %>% 
    tidyr::pivot_longer(-1) %>%
    tidyr::pivot_wider(names_from = 1, values_from = value)
}

df_meta 			<- data.frame(matrix(, nrow = 0, ncol = 10)) # dataframe to fill with data, with columns according to documentation of {metafor}-package
colnames(df_meta) 	<- c("study", "source", "comparator1", "comparator2", "n1i", "m1i", "sd1i", "n2i", "m2i", "sd2i")

df_doublecheck 		<- data.frame(matrix(, nrow = 0, ncol = 10)) # dataframe for studies which showed some sort of problem
colnames(df_doublecheck) <- c("study", "source", "comparator1", "comparator2", "n1i", "m1i", "sd1i", "n2i", "m2i", "sd2i")

for (i in 1:dim(idx_outcomes_rct1)[1]){ # loop through all available studies and fill df_meta
idx_total_list 		<- as.numeric(idx_outcomes_rct1$type[i]) # index of patient in the entire list (cf. first_authors)
results_temp 		<- extracted_results1_p[[i]] # the results of the primary outcomes
df_temp 			<- data.frame(matrix(, nrow = 1, ncol = 10))
colnames(df_temp) 	<- c("study", "source", "comparator1", "comparator2", "n1i", "m1i", "sd1i", "n2i", "m2i", "sd2i")

df_temp$study = i
df_temp$source <- idx_outcomes_rct1$author[i]

data_temp <- idx_outcomes_rct1 %>% 
	select(matches("outcome")) %>% 
	slice(i) %>% 
	t()

colnames(data_temp) <- "outcome"
columns_of_interest <- data.frame(data_temp) %>% filter(rowSums(across(everything(), ~grepl("follow-up", .x))) > 0)
if (dim(columns_of_interest)[1] > 2) {	
	print(sprintf("Processing study: %s", idx_outcomes_rct1$author[i]))
	df_doublecheck <- rbind(df_doublecheck, df_temp)
	# next
	
	# loop through all possible combinations and append all available data
	all_effects 				<- combn(columns_of_interest$outcome, 2)
	extracted_results2_numbers 	<- lapply(tab_names[as.numeric(idx_outcomes_rct1$type[i])], function(x) read_excel(results_file, 
															sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31"))
	df_numbers 					<- df_transpose(extracted_results2_numbers[[1]])
	colnames(df_numbers)		<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')
	df_temp_multiple 			<- df_temp # necessary to maintain the structure for all combinations
	
	for (k in 1:dim(all_effects)[2]){
		df_temp_multiple <- df_temp # necessary to maintain the structure for all combinations
		idx_comparator1 <- which(columns_of_interest$outcome==all_effects[,k][1])
		idx_comparator2 <- which(columns_of_interest$outcome==all_effects[,k][2])
		
		df_temp_multiple$comparator1 <- str_split(columns_of_interest$outcome[idx_comparator1], '_')[[1]][1]
		idx_numbers 		<- which(df_numbers$intervention==df_temp_multiple$comparator1) # column in which data is saved
		if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 
		temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator1, '_', 'follow-up'), .x))) > 0)
		df_temp_multiple$m1i 		<- as.numeric(temp_vals$mean)
		df_temp_multiple$sd1i 		<- as.numeric(temp_vals$sd)
		df_temp_multiple$n1i 		<- as.numeric(df_numbers$included[idx_numbers])
			
		df_temp_multiple$comparator2 <- str_split(columns_of_interest$outcome[idx_comparator2], '_')[[1]][1]
		idx_numbers 		<- which(df_numbers$intervention==df_temp_multiple$comparator2) # column in which data is saved
		if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 
		temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp_multiple$comparator2, '_', 'follow-up'), .x))) > 0)
		df_temp_multiple$m2i 		<- as.numeric(temp_vals$mean)
		df_temp_multiple$sd2i 		<- as.numeric(temp_vals$sd)
		df_temp_multiple$n2i 		<- as.numeric(df_numbers$included[idx_numbers])
		
		df_meta <- rbind(df_meta, df_temp_multiple) # append to list
	}
	
	
} else if (dim(columns_of_interest)[1]==2) {
	print(sprintf("Processing study: %s", idx_outcomes_rct1$author[i]))
	
	extracted_results2_numbers 	<- lapply(tab_names[as.numeric(idx_outcomes_rct1$type[i])], function(x) read_excel(results_file, 
															sheet = x, col_names=c("description", paste0('arm', c(1:6))), range = "C20:I31"))
	df_numbers <- df_transpose(extracted_results2_numbers[[1]])
	colnames(df_numbers)<- c('arms', 'intervention', 'eligible', 'enrolled', 'included', 'excluded', 'female_perc', 'age_mean', 'age_sd', 'age_mdn', 'hy', 'updrs_mean', 'updrs_sd')

	df_temp$comparator1 <- str_split(columns_of_interest$outcome, '_')[[1]][1]
	idx_numbers 		<- which(df_numbers$intervention==df_temp$comparator1) # column in which data is saved
	if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 
	temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp$comparator1, '_', 'follow-up'), .x))) > 0)
	df_temp$m1i 		<- as.numeric(temp_vals$mean)
	df_temp$sd1i 		<- as.numeric(temp_vals$sd)
	df_temp$n1i 		<- as.numeric(df_numbers$included[idx_numbers])
		
	df_temp$comparator2 <- str_split(columns_of_interest$outcome, '_')[[2]][1]
	idx_numbers 		<- which(df_numbers$intervention==df_temp$comparator2) # column in which data is saved
	if (identical(idx_numbers, integer(0))){ idx_numbers <- which(df_numbers$intervention=='total')} 
	temp_vals 			<- results_temp %>% filter(rowSums(across(everything(), ~grepl(paste0(df_temp$comparator2, '_', 'follow-up'), .x))) > 0)
	df_temp$m2i 		<- as.numeric(temp_vals$mean)
	df_temp$sd2i 		<- as.numeric(temp_vals$sd)
	df_temp$n2i 		<- as.numeric(df_numbers$included[idx_numbers])
	
	df_meta <- rbind(df_meta, df_temp) # append to list
} else {
	print(sprintf("There is a problem with study: %s", idx_outcomes_rct1$author[i]))
	#next
	}

}

df_test_meta_CTRL <- df_meta_CTRL.md %>% drop_na(m_pre, m_post, sd_pre, sd_post)
df_test_meta_CTRL$m_pre <- df_test_meta_CTRL$m_pre * (-1)
df_test_meta_TRT <- df_meta_TRT.md %>% drop_na(m_pre, m_post, sd_pre, sd_post)
df_test_meta_TRT$m_pre <- df_test_meta_TRT$m_pre * (-1)

df_test_meta_CTRL$ri <- 0
df_test_meta_TRT$ri <- 0
datC <- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_test_meta_CTRL)
datT <- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_test_meta_TRT)
dat <- data.frame(yi = datT$yi - datC$yi, vi = datT$vi + datC$vi, slab=datC$source)
dat1 <- dat
comparator1 <- cbind(df_test_meta_TRT$comparator, df_test_meta_CTRL$comparator)

df_test_meta_CTRL <- df_meta_CTRL %>% drop_na(m_pre, m_post, sd_pre, sd_post)
df_test_meta_TRT <- df_meta_TRT %>% drop_na(m_pre, m_post, sd_pre, sd_post)

df_test_meta_CTRL$ri <- 0
df_test_meta_TRT$ri <- 0
datC <- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_test_meta_CTRL)
datT <- escalc(measure="SMCR", m1i=m_post, m2i=m_pre, sd1i=sd_pre, ni=ni, ri=ri, slab=source, data=df_test_meta_TRT)
dat <- data.frame(yi = datT$yi - datC$yi, vi = datT$vi + datC$vi, slab=datC$source)
dat2 <- dat
comparator2 <- cbind(df_test_meta_TRT$comparator, df_test_meta_CTRL$comparator)

dat_all <- rbind(dat1, dat2)

res <- rma(yi, vi, data=dat_all, method="EE", digits=2)
res$slab <- dat_all$slab

### forest plot with extra annotations
forest(res, atransf=exp, at=log(c(.05, .25, 1, 4, 20)), xlim=c(-16,6),
       ilab=rbind(comparator1, comparator2),
       ilab.xpos=c(-9.5,-6.5), cex=.75, header="Author(s) and Year",
       mlab="")
op <- par(cex=.75, font=2)
text(c(-9.5,-6.5), dim(dat)[1]+2, c("Treatment", "Control"))
#text(c(-8.75,-5.25),     16, c("Vaccinated", "Control"))
par(op)


# TO BE CHECKED FOR:
idx <- 1:dim(idx_outcomes_rct1)[1]
idx <- idx[-c(1:12, 14, 16, 17, 19, 20)]
idx_outcomes_rct1$author[idx]

							  
### TEST

test_data <- data.frame(
m_pre   = c(23.1, 24.9, 0.6, 55.7, 34.8, -.2),
m_post  = c(19.7, 25.3, 0.6, 60.7, 33.4, 0),
sd_pre  = c(13.8, 4.1, 0.2, 17.3, 3.1, 2.1),
sd_post = c(14.8, 3.3, 0.2, 17.9, 6.9, 0),
ni      = c(20, 42, 9, 11, 14, 30),
ri      = c(.47, .64, .77, .89, .44, 0))

							  
dat2 <- escalc(measure="SMCR", m1i=m_pre, sd1i=sd_pre, ni=ni,
                              m2i=m_post, sd2i=sd_post, ri=ri, data=test_data)
dat3 <- escalc(measure="SMCC", m1i=m_pre, sd1i=sd_pre, ni=ni,
                              m2i=m_post, sd2i=sd_post, ri=ri, data=test_data)