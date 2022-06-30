	# This is code to run meta analyses on the results from the systematic literature review
	# For further details on the data extraction cf. worksheets in __included.xlsx and 
	# code in analysis_dataframe.r 
	# Code developed by David Pedrosa

	# Version 1.3 # 2022-06-22, refined analyses and added some explanations

	# ==================================================================================================
	## In case of multiple people working on one project, this helps to create an automatic script
	username = Sys.info()["login"]
	if (username == "dpedr") {
		wdir = "D:/Jottacloud/PDtremor_syst-review/"
		windowsFonts(Arial=windowsFont("TT Arial"))
	} else if (username == "david") {
		wdir = "/media/storage/Jottacloud/PDtremor_syst-review/"
	}
	setwd(wdir)

	# ==================================================================================================
	## Specify packages of interest and load them automatically if needed
	packages = c("readxl", "dplyr", "plyr", "tibble", "tidyr", "stringr", "openxlsx", "rstatix", "ggpubr",
					"metafor", "tidyverse", "clubSandwich", "multcomp", "reshape2", "latex2exp") # packages needed

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
	# Functions needed for plotting (taken from {metafor}-documentation)
	mlabfun <- function(text, res) { # helper function to add Q-test, I^2, and tau^2 estimate info
	   list(bquote(paste(.(text),
		  " (Q = ", .(formatC(res$QE, digits=2, format="f")),
		  ", df = ", .(res$k - res$p),
		  ", p ", .(metafor:::.pval(res$QEp, digits=2, showeq=TRUE, sep=" ")), "; ",
		  I^2, " = ", .(formatC(res$I2, digits=1, format="f")), "%, ",
		  tau^2, " = ", .(formatC(res$tau2, digits=2, format="f")), ")")))}
	# =================================================================================================
	
	# Start analyses:
	# ==================================================================================================
	# A. Three-level hierarchical model including all studies for effect sizes and category as 2nd level
	# ==================================================================================================

	# Prepare data adding groups, an "order" and a category
	# =================
	dat_results <- read.csv(file.path(wdir, "results", "results_meta-analysis.csv"), # load extracted data 
							stringsAsFactors = F,  encoding="UTF-8")

	groups 		<- data.frame(	levodopa=c("levodopa", rep(NA, 8)), # adds groups to dataframe
								betablocker=c("propanolol", "zuranolone", "nipradilol", rep(NA, 6)),
								primidone=c("primidone", rep(NA, 8)),
								anticholinergics=c("trihexphenidyl", "benztropin", "clozapine", rep(NA,6 )),
							  	# anticholinergics=c("trihexphenidyl", "benztropin", rep(NA,7 )),
								amantadine=c("amantadine", rep(NA, 8)), 
								dopamine_agonists=c("ropinirole", "rotigotine", "apomorphine", "pramipexole", 
													"piribedil", "bromocriptine", "pergolide", "cabergoline", 
													"dihydroergocriptin"), 
								# clozapine=c("clozapine", rep(NA, 8)), ## was added to anticholinergics 
								mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)),
								cannabis=c("cannabis", rep(NA, 8)), 
								budipine=c("budipine", rep(NA, 8)),
								zonisamide=c("zonisamide", rep(NA, 8)),
								adenosine_antagonist=c("theophylline", "caffeine", "adenosineA2a", rep(NA, 6)),
								botox=c("botulinum", rep(NA, 8)))

	dat 	<- dat_results %>% mutate(category=NA, class=NA, order_no=NA, treatment_diff=NA) # add columns to work with
	iter	<- 0
	for (k in colnames(groups)){ # assigns category to the different treatments
		iter 					<- iter + 1
		temp 					<- dat_results %>% rowid_to_column %>% 
									filter(if_any(everything(),~str_detect(., c(paste(groups[[k]][!is.na(groups[[k]])],collapse='|'))))) 
		dat$category[temp$rowid]<- colnames(groups)[iter]
		dat$order_no[temp$rowid]<- iter
	}

	sorted_indices 	<- dat %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
								count() %>% arrange(desc(order_no))# 

	dat 				<- dat %>% drop_na(-c(class, treatment_diff)) # remove NA, except to those in class and treatment_diff columns
	dat$study 			<- 1:dim(dat)[1] # add study ID in the data frame for estimating results later
	dat 				<- dat %>% arrange(desc(order_no), treatment)
	attr(dat$yi, 'slab')<- dat$slab			
	temp 				<- dat %>% rowid_to_column %>% filter(if_any(everything(),
										~str_detect(., c(paste(c("levodopa", "amantadine", "clozapine", "budipine", "cannabis", 
																	"zonisamide", "primidone", "botox"), collapse='|'))))) 
	dat$treatment_diff[setdiff(1:dim(dat)[1], temp$rowid)] = dat$treatment[setdiff(1:dim(dat)[1], temp$rowid)]

	coal = TRUE
	if (coal==TRUE) { # merges columns from "treatment" and "treatment_diff" where NA appear 
		dat <- dat %>% mutate(treatment_diff = coalesce(treatment_diff, treatment))
	}

	# Prepare contingency table for quality of life estimates
	# =================
	dat <- dat %>% mutate(., qualsyst_cat = case_when(
								 qualsyst <= .50 ~ "low",
								 (qualsyst > .50 & qualsyst < .70) ~ "medium",
								 qualsyst >= .70 ~ "high"))

	# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
	# =================
	xrows_temp 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()

	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	
	# Impute Variance Covariance matrix according to correlated effects (see above)
	# =================
	covariance_adaptation = TRUE 		# in case of studies with the same participants, covariance should be adopted cf. 
	if (covariance_adaptation==TRUE) { 	# https://www.jepusto.com/imputing-covariance-matrices-for-multi-variate-meta-analysis/
		dat$study.id <- dat$study
		dat$study.id[which(dat$slab=="Koller et al. 1987.1")] <- dat$study[which(dat$slab=="Koller et al. 1987")]
		dat$study.id[which(dat$slab=="Brannan & Yahr (1995)")] <- dat$study[which(dat$slab=="Brannan & Yahr (1995)")][1]
		dat$study.id[which(dat$slab=="Sahoo et al. 2020")] <- dat$study[which(dat$slab=="Sahoo et al. 2020")][1]
		dat$study.id[which(dat$slab=="Navan et al. 2003.1")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][1]
		dat$study.id[which(dat$slab=="Navan et al. 2003")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][2]
		dat$study.id[which(dat$slab=="Friedman et al. 1997")] <- dat$study[which(dat$slab=="Friedman et al. 1997")][1]
		dat$study.id[which(dat$slab=="Zach et al. 2020")] <- dat$study[which(dat$slab=="Dirkx et al. 2019")][1]
	}

	V_mat <- impute_covariance_matrix(vi = dat$vi, cluster=dat$study, r = .7)
	
	# Run random effects meta-analysis using some "moderators" (qualsyst scores and study_type)
	# =================	
	res <- rma.mv(yi, slab=dat$slab, V=V_mat,
					random = ~ 1 | as.factor(category)/study, 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG", #qualsyst*study_type
					data=dat, verbose=TRUE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
	res_wo1 <- rma.mv(yi, slab=dat$slab, V=V_mat,
					random = ~1 | as.factor(category)/study, 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG", #qualsyst*study_type
					data=dat, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(0, NA))
	res_wo2 <- rma.mv(yi, slab=dat$slab, V=V_mat,
					random = ~1 | as.factor(category)/study, 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG", #qualsyst*study_type
					data=dat, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(NA, 0))
	comp_allstudies1 <- anova(res, res_wo1)
	comp_allstudies2 <- anova(res, res_wo1)

	# Run model diagnostics
	# =================	
	# Export diagnostics (profile) on model with all studies/categories	
	svg(filename=file.path(wdir, "results", "supplFigure.profile_all_groups.v1.0.svg"))
	par(mfrow=c(2,1))
	profile(res, sigma2=1, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{1}$ (factor = category))'))
	title("Profile Likelihood Plot for model including all subgroups", line = -1, outer = TRUE)
	profile(res, sigma2=2, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{2}$ (factor = category/study_type))'))
	dev.off()

	# Calculation of I^2 statistic
	Wtot <- diag(1/res$vi)
	Xtot <- model.matrix(res)
	Ptot <- Wtot - Wtot %*% Xtot %*% solve(t(Xtot) %*% Wtot %*% Xtot) %*% t(Xtot) %*% Wtot
	100 * sum(res$sigma2) + (sum(res$sigma2) + (res$k-res$p)/sum(diag(Ptot)))
	100 * res$sigma2 / (sum(res$sigma2) + (res$k-res$p)/sum(diag(Ptot)))

	# Start plotting data using forest plot routines from {metafor}-package
	# =================	
	cex_plot 	<- 1 		# plot size
	y_lim 		<- start 	# use last "start" as y-limit for forest plot
	y_offset 	<- .5		# offset used to plot the header of the column
	svg(filename=file.path(wdir, "results", "meta_analysis.all_groups.v2.1.svg"))

	if (Sys.info()['sysname']!="Linux") {
		windowsFonts("Arial" = windowsFont("TT Arial"))
	}
	
	forest(res, xlim=c(-10, 4.6), at=c(-3, -2, -1, 0, 1, 2), #atransf=exp,
		   ilab=cbind(dat$ni, dat$treatment_diff, dat$study_type, 
		   sprintf("%.02f",  dat$qualsyst), 
		   paste0(formatC(weights(res), format="f", digits=1, width=4), "%")),
		   xlab="Standardised mean change [SMCR]",
		   ilab.xpos=c(-5, -4, 2.5, 3, 3.5), 
		   cex=cex_plot, ylim=c(-1, y_lim),
		   rows=xrows,
		   efac =c(.8), #offset=y_offset, 
		   mlab=mlabfun("RE Model for All Studies", res),
		   font=4, header="Author(s) and Year")
	 
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(font=2, mar = c(2, 2, 2, 2))

	text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
			pos=c(2, 4), offset=y_offset)

	### Add column headings for everything defined in "forest::ilab"
	text(c(-5), cex=cex_plot, y_lim-y_offset, c("n"))
	text(c(-4), cex=cex_plot, y_lim-y_offset, c("agent"))
	text(c(2.5), cex=cex_plot, y_lim-y_offset, c("type"))
	text(c(3), cex=cex_plot, y_lim+3*y_offset, c("QualSyst"))
	text(c(3), cex=cex_plot, y_lim-y_offset, c("score"))
	text(c(3.5), cex=cex_plot, y_lim-y_offset, c("weight"))
	 
	### Switch to bold-italic font at normal size for plotting the subgroup headers
	par(cex=cex_plot, font=4)
	 
	### add text for the subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.5, c(sprintf("Levodopa (n = %s)", sum(dat$ni[dat$category=="levodopa"])),
								   sprintf("beta-blocker (n = %s)", sum(dat$ni[dat$category=="betablocker"])),
								   sprintf("Primidone (n = %s)", sum(dat$ni[dat$category=="primidone"])),
								   sprintf("Anticholinergics (n = %s)", sum(dat$ni[dat$category=="anticholinergics"])),
								   sprintf("Amantadine (n = %s)", sum(dat$ni[dat$category=="amantadine"])),
								   sprintf("Dopamine agonists (n = %s)", sum(dat$ni[dat$category=="dopamine_agonists"])), 
								   #sprintf("Clozapine (n = %s)", sum(dat$ni[dat$category=="clozapine"])), 
								   sprintf("MAO-inhibitor (n = %s)", sum(dat$ni[dat$category=="mao_inhibitors"])), 
								   sprintf("Cannabis (n = %s)", sum(dat$ni[dat$category=="cannabis"])),
								   sprintf("Budipine (n = %s)", sum(dat$ni[dat$category=="budipine"])),				
								   sprintf("Zonisamide (n = %s)", sum(dat$ni[dat$category=="zonisamide"])),
								   sprintf("Adenosine antagonists (n = %s)", sum(dat$ni[dat$category=="adenosine_antagonist"])),
								   sprintf("Botulinum toxin (n = %s)", sum(dat$ni[dat$category=="botox"]))))

	### set par back to the original settings
	par(op)
	 
	### fit random-effects model in all subgroups
	res.lev <- rma(yi, vi, subset=(category=="levodopa"), data=dat)
	res.bet <- rma(yi, vi, subset=(category=="betablocker"), data=dat)
	res.pri <- rma(yi, vi, subset=(category=="primidone"), data=dat)
	res.cho <- rma(yi, vi, subset=(category=="anticholinergics"), data=dat)
	res.ama <- rma(yi, vi, subset=(category=="amantadine"),     data=dat)
	res.dag <- rma(yi, vi, subset=(category=="dopamine_agonists"),  data=dat)
	# res.clo <- rma(yi, vi, subset=(category=="clozapine"),  data=dat) ## was added to anticholinergics
	res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"),     data=dat)
	res.cbd <- rma(yi, vi, subset=(category=="cannabis"),  data=dat)
	res.bud <- rma(yi, vi, subset=(category=="budipine"),  data=dat)
	res.zon <- rma(yi, vi, subset=(category=="zonisamide"),  data=dat)
	res.ade <- rma(yi, vi, subset=(category=="adenosine_antagonist"),  data=dat)
	res.bot <- rma(yi, vi, subset=(category=="botox"),  data=dat)

	par(font=4)

	yshift = .2
	fac_cex = .8

	if (Sys.info()['sysname']!="Linux") {
		windowsFonts("Arial" = windowsFont("TT Arial"))
	}

	### add summary polygons for the three subgroups
	addpoly(res.lev, row= 114.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.bet, row= 107.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pri, row= 102.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pri), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.cho, row= 91.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.ama, row= 85.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ama), cex=cex_plot*fac_cex, efac = c(.5))
	# addpoly(res.clo, row= 60.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.dag, row= 51.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.dag), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.mao, row= 35.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.cbd, row= 30.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.cbd), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.bud, row= 22.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bud), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.zon, row= 16.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.zon), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.ade, row= 9.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ade), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.bot, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bot), cex=cex_plot*fac_cex, efac = c(.5))
	 
	abline(h=0)

	### fit meta-regression model to test for subgroup differences
	res_sg <- rma(yi, vi, mods = ~ category, data=dat)
	 
	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_sg$QM, digits=2, format="f")), ", df = ", .(res_sg$p - 1),
		 ", p = ", .(formatC(res_sg$QMp, digits=3, format="f")))))
	# addpoly(res_sg, row= -1.8, mlab=mlabfun("RE Model for Subgroup", res_sg), cex=cex_plot*fac_cex, efac = c(.5))
	
	if (Sys.info()['sysname']!="Linux") {
		windowsFonts("Arial" = windowsFont("TT Arial"))
	}
	dev.off()
		
	# Run model diagnostics
	# =================	
	# Estimate bias according to funnel plot and Egger's statistics
	svg(filename=file.path(wdir, "results", "funnel.all_groups.v1.0.svg"))
	funnel_contour <- funnel(res, level=c(90, 95, 99), back="gray95",
                  shade=c("white", "gray75", "gray85", "gray95"), 
                  refline=0, legend="topleft", ylim=c(0, .8), las=1, digits=list(1L, 1), 
				  xlab="Standardized mean change")
	dev.off() 
	
	dat$residuals	<- residuals.rma(res)
	dat$precision	<- 1/sqrt(dat$vi)
	egger_res.all	<- rma.mv(residuals~precision,vi,data=dat, random = ~ factor(study) | factor(category))
		
	# ==================================================================================================
	# B. Two-level hierarchical model with studies using dopamine agonists
	# ==================================================================================================

	# Prepare data adding groups, an "order" and a category
	# =================
	dat_dopamineTotal 		<- dat %>% filter(category %in% c("levodopa", "dopamine_agonists"))	
	remove_ergot_derivates=TRUE
	if (remove_ergot_derivates){ # removes ergot derivatives as they are not marketed anymore
		dat_dopamineTotal 	<- dat_dopamineTotal %>% filter(treatment_diff!="dihydroergocriptin") %>% 
			filter(treatment_diff!="cabergoline") %>% 
			filter(treatment_diff!="pergolide")
	}
	
	groups 	<- c("levodopa", "rotigotin", "ropinirole", "pramipexole", "piribedil", "apomorphine")
	iter	<- 0
	for (i in groups){ 
		iter <- iter + 1
		temp <- dat_dopamineTotal %>% rowid_to_column %>% 
			filter(if_any(everything(),~str_detect(., i))) 
		dat_dopamineTotal$order_no[temp$rowid]<- iter
	}
	
	sorted_indices 	<- dat_dopamineTotal %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 

	# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
	# =================
	xrows_temp 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()

	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	# Impute Variance Covariance matrix according to correlated effects (see above)
	# =================
	V_mat_agonists <- impute_covariance_matrix(vi = dat_dopamineTotal$vi, cluster=dat_dopamineTotal$study, r = .7)

	# Run random effects meta-analysis using some "moderators" (qualsyst scores and study_type)
	# =================	
	res_agonistsTotal 	<- rma.mv(yi, slab=dat_dopamineTotal$slab, V=V_mat_agonists,
						random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
						mods= ~ qualsyst*study_type, tdist=TRUE, #struct="DIAG",
						data=dat_dopamineTotal, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
	res_agonistsTotalwo	<- rma.mv(yi, slab=dat_dopamineTotal$slab, V=V_mat_agonists,
						random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
						mods= ~ qualsyst*study_type, tdist=TRUE, #struct="DIAG",
						data=dat_dopamineTotal, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(0, NA))
	comp_agonistsRCT 	<- anova(res_agonistsTotal, res_agonistsTotalwo)
	
	# Run model diagnostics
	# =================	
	# Export diagnostics (profile) on model with all studies/categories	
	svg(filename=file.path(wdir, "results", "supplFigure.profile_all_agonists.v1.0.svg"))
	par(mfrow=c(2,1))
	profile(res_agonistsTotal, sigma2=1, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{1}$ (factor = category))'), progbar=TRUE)
	title("Profile Likelihood Plot for model including only levodopa and dopamine agonists", line = -1, outer = TRUE)
	profile(res_agonistsTotal, sigma2=2, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{2}$ (factor = category/study_type))'), progbar=TRUE)
	dev.off()

	# Calculation of I^2 statistic
	Wago <- diag(1/res_agonistsTotal$vi)
	Xago <- model.matrix(res_agonistsTotal)
	Pago <- Wago - Wago %*% Xago %*% solve(t(Xago) %*% Wago %*% Xago) %*% t(Xago) %*% Wago
	100 * sum(res_agonistsTotal$sigma2) + (sum(res_agonistsTotal$sigma2) + (res_agonistsTotal$k-res_agonistsTotal$p)/sum(diag(Pago)))
	100 * res_agonistsTotal$sigma2 / (sum(res_agonistsTotal$sigma2) + (res_agonistsTotal$k-res_agonistsTotal$p)/sum(diag(Pago)))
	
	# Start plotting data using forest plot routines from {metafor}-package
	# =================	
	cex_plot 	<- 1 		# plot size
	y_lim 		<- start 	# use last "start" as y-limit for forest plot
	y_offset 	<- .5		# offset used to plot the header of the column
	svg(filename=file.path(wdir, "results", "meta_analysis.agonists_levodopa.v2.1.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	forest(res_agonistsTotal, xlim=c(-10, 4.6), #at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
			ilab=cbind(dat_dopamineTotal$ni, dat_dopamineTotal$study_type, 
			sprintf("%.02f",  dat_dopamineTotal$qualsyst), 
			paste0(formatC(weights(res_agonistsTotal), format="f", digits=1, width=4), "%")),
			xlab="Standardised mean change [SMCR]",
			ilab.xpos=c(-5, 2.5, 3, 3.5), 
			cex=cex_plot, ylim=c(-1, y_lim),
			rows=xrows,
			efac =c(.8), #offset=y_offset, 
			mlab=mlabfun("RE Model for All Studies", res_agonistsTotal),
			font=4, header="Author(s) and Year")
		   
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(font=2, mar = c(2, 2, 2, 2))

	text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
			pos=c(2, 4), offset=y_offset)

	### Add column headings for everything defined in "forest::ilab"
	text(c(-5), cex=cex_plot, y_lim-y_offset, c("n"))
	# text(c(-4), cex=cex_plot, y_lim-y_offset, c("agent"))
	text(c(2.5), cex=cex_plot, y_lim-y_offset, c("type"))
	text(c(3), cex=cex_plot, y_lim+3*y_offset, c("QualSyst"))
	text(c(3), cex=cex_plot, y_lim-y_offset, c("score"))
	text(c(3.5), cex=cex_plot, y_lim-y_offset, c("weight"))
		
	### Switch to bold italic font
	par(font=4)

	### Adds text for subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.1, c(sprintf("Levodopa (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="levodopa"])),
								   sprintf("Rotigotine (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="rotigotine"])),
								   sprintf("Ropinirole (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="ropinirole"])),
								   sprintf("Pramipexole (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="pramipexole"])),
								   sprintf("Piribedil (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="piribedil"])),
								   sprintf("Apomorphine (n = %s)", sum(dat_dopamineTotal$ni[dat_dopamineTotal$treatment_diff=="apomorphine"]))))

	### set par back to the original settings
	par(op)

	### fit random-effects model in all subgroups
	res.apo <- rma(yi, vi, subset=(treatment_diff=="apomorphine"), data=dat_dopamineTotal)
	# res.cab <- rma(yi, vi, subset=(treatment2=="cabergoline"), data=dat_agonists)
	# res.per <- rma(yi, vi, subset=(treatment2=="pergolide"), data=dat_agonists)
	# res.dih <- rma(yi, vi, subset=(treatment2=="dihydroergocriptin"), data=dat_agonists)
	res.rot <- rma(yi, vi, subset=(treatment_diff=="rotigotine"), data=dat_dopamineTotal)
	res.rop <- rma(yi, vi, subset=(treatment_diff=="ropinirole"), data=dat_dopamineTotal)
	res.pir <- rma(yi, vi, subset=(treatment_diff=="piribedil"), data=dat_dopamineTotal)
	res.pra <- rma(yi, vi, subset=(treatment_diff=="pramipexole"), data=dat_dopamineTotal)
	res.lev <- rma(yi, vi, subset=(treatment_diff=="levodopa"), data=dat_dopamineTotal)

	yshift = .2
	fac_cex = .75
	### add summary polygons for the three subgroups
	addpoly(res.apo, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.apo), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pir, row= 9.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pir), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pra, row= 19.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pra), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.rop, row= 31.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rop), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.rot, row= 39.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rot), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.lev, row= 46.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*fac_cex, efac = c(.5))
	
	abline(h=0)
	
	### fit meta-regression model to test for subgroup differences
	dat_dopamineTotal$treatment_diff <- as.factor(dat_dopamineTotal$treatment_diff)
	groups <- c("levodopa", "apomorphine", "piribedil", "pramipexole", "ropinirole", "rotigotine")
	levels(dat_dopamineTotal) <- groups
	res_overall_agonistsTotal <- rma(yi, vi, mods = ~ treatment_diff - 1, data=dat_dopamineTotal)

	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overall_agonistsTotal$QM, digits=2, format="f")), ", df = ", .(res_overall_agonistsTotal$p - 1),
		 ", p = ", .(formatC(res_overall_agonistsTotal$QMp, digits=3, format="f")))))
	dev.off()
	
	# Run post-hoc analyses
	# =================	
	pwc 					<- summary(glht(res_overall_agonistsTotal, linfct=cbind(contrMat(rep(1,6), type="Tukey"))), 
							test=adjusted("BH")) # pairwise comparisons
	mat2plot 				<- matrix(0, nrow = 6, ncol = 6, dimnames=list(toupper(substr(groups, 1, 3)), toupper(substr(groups, 1, 3)))) + + diag(6)/1000000
	mat_significance 		<- mat2plot 
	mat2plot[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$tstat)
	mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
	melted_cormat 			<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
	melted_cormat$pval 		<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
	melted_cormat$pval 		<- melted_cormat$pval$value # dodgy solution
	colnames(melted_cormat)	<-c("Var1", "Var2", "tstat", "pval")
	melted_cormat[melted_cormat==0] <- NA # remove zero values and convert to NA
	melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001] <- "*"
	melted_cormat$pval[melted_cormat$pval<.001] <- "**"
	melted_cormat$pval[melted_cormat$pval>.05] <- NA

	# Create a heatmap for pairwise comparisons
	svg(filename=file.path(wdir, "results", "heatmaps.agonists_levodopa.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	ggplot(data = melted_cormat %>% drop_na(-pval), aes(Var2, Var1, fill = tstat))+
		geom_tile(color = "white")+
		scale_fill_gradient2(mid="#FBFEF9",low="#0C6291",high="#A63446", limits=c(-5,5),space = "Lab", 
						   name="T-statistics") +
		geom_text(aes(label = pval)) +
		ylab(NULL)  + xlab(NULL) +
		ggtitle("Pairwise comparison between levodopa and dopamine agonists") + 
		theme_minimal() +
		#theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   #size = 12, hjust = 1), 
									   #axis.text.x=element_blank()) +
		coord_fixed()
	dev.off()
	
	# ==================================================================================================
	# ==================================================================================================
	# C. Two-level hierarchical model with studies using dopamine agonists (only RCTs)
	# ==================================================================================================
	# ==================================================================================================

	# Prepare data adding groups, an "order" and a category
	# =================
	dat_agonistsRCT <- dat %>% filter(category %in% c("levodopa", "dopamine_agonists")) %>% filter(study_type=="RCT")
	remove_ergot_derivates=TRUE
	if (remove_ergot_derivates){ # removes ergot derivatives as they are not marketed anymore
		dat_agonistsRCT 	<- dat_agonistsRCT %>% filter(treatment_diff!="dihydroergocriptin") %>% 
			filter(treatment_diff!="cabergoline") %>% 
			filter(treatment_diff!="pergolide")
	}
	
	groups 	<- c("levodopa", "rotigotin", "ropinirole", "pramipexole", "piribedil", "apomorphine")
	iter	<- 0
	for (i in groups){ 
		iter <- iter + 1
		temp <- dat_agonistsRCT %>% rowid_to_column %>% 
			filter(if_any(everything(),~str_detect(., i))) 
		dat_agonistsRCT$order_no[temp$rowid]<- iter
	}
	
	sorted_indices 	<- dat_agonistsRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 

	# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
	# =================
	xrows_temp 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()

	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	# Impute Variance Covariance matrix according to correlated effects (see above)
	# =================
	V_mat_agonistsRCT 	<- impute_covariance_matrix(vi = dat_agonistsRCT$vi, cluster=dat_agonistsRCT$study, r = .7)

	# Run random effects meta-analysis using moderators (qualsyst scores), study_type obvioulsy not necessary anymore
	# =================	
	res_agonistsRCT	<- rma.mv(yi, slab=dat_agonistsRCT$slab, V=V_mat_agonistsRCT,
					random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_agonistsRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))

	# Run model diagnostics
	# =================	
	# Export diagnostics (profile) on model with all studies/categories	
	svg(filename=file.path(wdir, "results", "supplFigure.profile_agonistsRCT.v1.0.svg"))
	par(mfrow=c(2,1))
	profile(res_agonistsRCT, sigma2=1, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{1}$ (factor = category))'), progbar=TRUE,)
	title("Profile Likelihood Plot for model including dopamine agonists at RCT", line = -1, outer = TRUE)
	profile(res_agonistsRCT, sigma2=2, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{2}$ (factor = category/study_type))'), progbar=TRUE,)
	dev.off()

	# Calculation of I^2 statistic
	WagoRCT <- diag(1/res_agonistsRCT$vi)
	XagoRCT <- model.matrix(res_agonistsRCT)
	PagoRCT <- WagoRCT - WagoRCT %*% XagoRCT %*% solve(t(XagoRCT) %*% WagoRCT %*% XagoRCT) %*% t(XagoRCT) %*% WagoRCT
	100 * sum(res_agonistsRCT$sigma2) + (sum(res_agonistsRCT$sigma2) + (res_agonistsRCT$k-res_agonistsRCT$p)/sum(diag(Pago)))
	100 * res_agonistsRCT$sigma2 / (sum(res_agonistsRCT$sigma2) + (res_agonistsRCT$k-res_agonistsRCT$p)/sum(diag(Pago)))

	# Start plotting data using forest plot routines from {metafor}-package
	# =================	
	cex_plot 	<- 1 		# plot size
	y_lim 		<- start 	# use last "start" as y-limit for forest plot
	y_offset 	<- .5		# offset used to plot the header of the column
	svg(filename=file.path(wdir, "results", "meta_analysis.agonistsRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	forest(res_agonistsRCT, xlim=c(-10, 5.6), at=c(-3, -2, -1, 0, 1, 2), #at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
			ilab=cbind(dat_agonistsRCT$ni, dat_agonistsRCT$study_type, 
			sprintf("%.02f",  dat_agonistsRCT$qualsyst), 
			paste0(formatC(weights(res_agonistsRCT), format="f", digits=1, width=4), "%")),
			xlab="Standardised mean change [SMCR]",
			ilab.xpos=c(-5, 2.5, 3, 3.5), 
			cex=cex_plot, ylim=c(-1, y_lim),
			rows=xrows,
			efac =c(.8), #offset=y_offset, 
			mlab=mlabfun("RE Model for All Studies", res_agonistsRCT),
			font=4, header="Author(s) and Year")

	
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(font=2, mar = c(2, 2, 2, 2))

	text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
			pos=c(2, 4), offset=y_offset)

	### Add column headings for everything defined in "forest::ilab"
	text(c(-5), cex=cex_plot, y_lim-y_offset, c("n"))
	# text(c(-4), cex=cex_plot, y_lim-y_offset, c("agent"))
	text(c(2.5), cex=cex_plot, y_lim-y_offset, c("type"))
	text(c(3), cex=cex_plot, y_lim+3*y_offset, c("QualSyst"))
	text(c(3), cex=cex_plot, y_lim-y_offset, c("score"))
	text(c(3.5), cex=cex_plot, y_lim-y_offset, c("weight"))
		
	### Switch to bold italic font
	par(font=4)

	### add text for subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.1, c(sprintf("Rotigotine (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment_diff=="rotigotine"])),
								   sprintf("Ropinirole (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment_diff=="ropinirole"])),
								   sprintf("Pramipexole (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment_diff=="pramipexole"])),
								   sprintf("Piribedil (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment_diff=="piribedil"])),
								   sprintf("Apomorphine (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment_diff=="apomorphine"]))))

	### set par back to the original settings
	par(op)
	
	### fit random-effects model in all subgroups
	res.apo <- rma(yi, vi, subset=(treatment_diff=="apomorphine"), data=dat_agonistsRCT)
	# res.cab <- rma(yi, vi, subset=(treatment2=="cabergoline"), data=dat_agonists)
	# res.per <- rma(yi, vi, subset=(treatment2=="pergolide"), data=dat_agonists)
	#res.dih <- rma(yi, vi, subset=(treatment2=="dihydroergocriptin"), data=dat_agonists)
	res.rot <- rma(yi, vi, subset=(treatment_diff=="rotigotine"), data=dat_agonistsRCT)
	res.rop <- rma(yi, vi, subset=(treatment_diff=="ropinirole"), data=dat_agonistsRCT)
	res.pir <- rma(yi, vi, subset=(treatment_diff=="piribedil"), data=dat_agonistsRCT)
	res.pra <- rma(yi, vi, subset=(treatment_diff=="pramipexole"), data=dat_agonistsRCT)
	#res.lev <- rma(yi, vi, subset=(treatment2=="levodopa"), data=dat_agonists)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.rot, row= 29.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rot), cex=cex_plot*fac_cex, efac = c(.5), font=3)
	addpoly(res.rop, row= 21.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rop), cex=cex_plot*fac_cex, efac = c(.5), font=3)
	addpoly(res.pra, row= 13.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pra), cex=cex_plot*fac_cex, efac = c(.5), font=3)
	addpoly(res.pir, row= 7.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pir), cex=cex_plot*fac_cex, efac = c(.5), font=3)
	addpoly(res.apo, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.apo), cex=cex_plot*fac_cex, efac = c(.5), font=3)
	par(op)
	abline(h=0)
	
	### fit meta-regression model to test for subgroup differences
	dat_agonistsRCT$treatment2 <- as.factor(dat_agonistsRCT$treatment_diff)
	groups <- c("apomorphine", "piribedil", "pramipexole", "ropinirole", "rotigotine")
	levels(dat_agonistsRCT) <- groups
	res_overall_agonistsRCT <- rma(yi, vi, mods = ~ treatment_diff - 1, data=dat_agonistsRCT)
	
	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overall_agonistsRCT$QM, digits=2, format="f")), ", df = ", .(res_overall_agonistsRCT$p - 1),
		 ", p = ", .(formatC(res_overall_agonistsRCT$QMp, digits=3, format="f")))))
	dev.off()

	# Run post-hoc analyses
	# =================	
	pwc 					<- summary(glht(res_overall_agonistsRCT, linfct=cbind(contrMat(rep(1,5), type="Tukey"))), 
							test=adjusted("BH")) # pairwise comparisons
	mat2plot 				<- matrix(0, nrow = 5, ncol = 5, dimnames=list(toupper(substr(groups, 1, 3)), toupper(substr(groups, 1, 3)))) + + diag(5)/1000000
	mat_significance 		<- mat2plot 
	mat2plot[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$tstat)
	mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
	melted_cormat 			<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
	melted_cormat$pval 		<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
	melted_cormat$pval 		<- melted_cormat$pval$value # dodgy solution
	colnames(melted_cormat)	<-c("Var1", "Var2", "tstat", "pval")
	melted_cormat[melted_cormat==0] <- NA # remove zero values and convert to NA
	melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001] <- "*"
	melted_cormat$pval[melted_cormat$pval<.001] <- "**"
	melted_cormat$pval[melted_cormat$pval>.05] <- NA

	# Create a heatmap for pairwise comparisons
	svg(filename=file.path(wdir, "results", "heatmaps.agonistsRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	ggplot(data = melted_cormat %>% drop_na(-pval), aes(Var2, Var1, fill = tstat))+
		geom_tile(color = "white")+
		scale_fill_gradient2(mid="#FBFEF9",low="#0C6291",high="#A63446", limits=c(-5,5),space = "Lab", 
						   name="T-statistics") +
		geom_text(aes(label = pval)) +
		ylab(NULL)  + xlab(NULL) +
		ggtitle("Pairwise comparison between dopamine agonists") + 
		theme_minimal() +
		#theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   #size = 12, hjust = 1), 
									   #axis.text.x=element_blank()) +
		coord_fixed()
	dev.off()


	# ==================================================================================================
	# ==================================================================================================
	# D. Two-level hierarchical model with studies using mao-inhibitors (only RCTs)
	# ==================================================================================================
	# ==================================================================================================

	# Prepare data adding groups, an "order" and a category
	# =================
	groups 		<- data.frame(	mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)))
	dat_maoRCT 	<- dat_results %>% mutate(category=NA, class=NA, order_no=NA, treatment_diff=NA) # add columns to dataframe to work with
	
	iter	<- 0
	for (k in colnames(groups)){ # assigns category to the different treatments
		iter 					<- iter + 1
		temp 					<- dat_results %>% rowid_to_column %>% 
									filter(if_any(everything(),~str_detect(., c(paste(groups[[k]][!is.na(groups[[k]])],collapse='|'))))) 
		dat_maoRCT$category[temp$rowid]<- colnames(groups)[iter]
		dat_maoRCT$order_no[temp$rowid]<- iter
	}

	sorted_indices 	<- dat_maoRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 

	dat_maoRCT 			<- dat_maoRCT %>% drop_na(-c(class, treatment_diff)) # get data of interest, that
	dat_maoRCT$study 	<- 1:dim(dat_maoRCT)[1] # add the study ID in the daat frame
	dat_maoRCT			<- dat_maoRCT %>% arrange(desc(order_no), treatment)
	attr(dat_maoRCT$yi, 'slab')<- dat_maoRCT$slab			
	temp 				<- dat_maoRCT %>% rowid_to_column %>% 
								filter(if_any(everything(),
										~str_detect(., c(paste(c("rasagiline", "selegiline", "safinamide")), collapse='|'))))) 
	dat_maoRCT$treatment_diff[setdiff(1:dim(dat_maoRCT)[1], temp$rowid)] = dat_maoRCT$treatment[setdiff(1:dim(dat_maoRCT)[1], temp$rowid)]

	coal = TRUE
	if (coal==TRUE) { # merges columns from "treatment" and "treatment_diff" where NA appear 
		dat_maoRCT <- dat_maoRCT %>% mutate(treatment_diff = coalesce(treatment_diff, treatment))
	}

	groups 	<- c("selegiline", "safinamide", "rasagiline")
	iter	<- 0
	for (i in groups){ 
		iter <- iter + 1
		temp <- dat_maoRCT %>% rowid_to_column %>% 
			filter(if_any(everything(),~str_detect(., i))) 
		dat_maoRCT$order_no[temp$rowid]<- iter
	}
	
	sorted_indices 	<- dat_maoRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 

	# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
	# =================
	xrows_temp 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()

	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	# Impute Variance Covariance matrix according to correlated effects (see above)
	# =================
	V_mat_maoRCT 	<- impute_covariance_matrix(vi = dat_maoRCT$vi, cluster=dat_maoRCT$study, r = .7)

	# Run random effects meta-analysis using moderators (qualsyst scores), study_type obviously not necessary anymore
	# =================	
	res_maoRCT	<- rma.mv(yi, slab=dat_maoRCT$slab, V=V_mat_maoRCT,
					random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_maoRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
	res_maoRCTwo<- rma.mv(yi, slab=dat_maoRCT$slab, V=V_mat_maoRCT,
					random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_maoRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(0, NA))


	# Run model diagnostics
	# =================	
	# Export diagnostics (profile) on model with all studies/categories	
	svg(filename=file.path(wdir, "results", "supplFigure.profile_maoRCT.v1.0.svg"))
	par(mfrow=c(2,1))
	profile(res_maoRCT, sigma2=1, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{1}$ (factor = category))'), progbar=TRUE,)
	title("Profile Likelihood Plot for model including MAO-inhibitors at RCT", line = -1, outer = TRUE)
	profile(res_maoRCT, sigma2=2, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{2}$ (factor = category/study_type))'), progbar=TRUE,)
	dev.off()

	# Calculation of I^2 statistic
	WmaoRCT <- diag(1/res_maoRCT$vi)
	XmaoRCT <- model.matrix(res_maoRCT)
	PmaoRCT <- WmaoRCT - WmaoRCT %*% XmaoRCT %*% solve(t(XmaoRCT) %*% WmaoRCT %*% XmaoRCT) %*% t(XmaoRCT) %*% WmaoRCT
	100 * sum(res_maoRCT$sigma2) + (sum(res_maoRCT$sigma2) + (res_maoRCT$k-res_maoRCT$p)/sum(diag(PmaoRCT)))
	100 * res_maoRCT$sigma2 / (sum(res_maoRCT$sigma2) + (res_maoRCT$k-res_maoRCT$p)/sum(diag(PmaoRCT)))

	# Start plotting data using forest plot routines from {metafor}-package
	# =================	
	cex_plot 	<- 1 		# plot size
	y_lim 		<- start 	# use last "start" as y-limit for forest plot
	y_offset 	<- .5		# offset used to plot the header of the column
	svg(filename=file.path(wdir, "results", "meta_analysis.maoRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	forest(res_maoRCT, xlim=c(-10, 5.6), at=c(-3, -2, -1, 0, 1, 2),#at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
			ilab=cbind(dat_maoRCT$ni, dat_maoRCT$study_type, 
			sprintf("%.02f",  dat_maoRCT$qualsyst), 
			paste0(formatC(weights(res_maoRCT), format="f", digits=1, width=4), "%")),
			xlab="Standardised mean change [SMCR]",
			ilab.xpos=c(-5, 2.5, 3, 3.5), 
			cex=cex_plot, ylim=c(-1, y_lim),
			rows=xrows,
			efac =c(.8), #offset=y_offset, 
			mlab=mlabfun("RE Model for All Studies", res_maoRCT),
			font=4, header="Author(s) and Year")

	
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(font=2, mar = c(2, 2, 2, 2))

	text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
			pos=c(2, 4), offset=y_offset)

	### Add column headings for everything defined in "forest::ilab"
	text(c(-5), cex=cex_plot, y_lim-y_offset, c("n"))
	# text(c(-4), cex=cex_plot, y_lim-y_offset, c("agent"))
	text(c(2.5), cex=cex_plot, y_lim-y_offset, c("type"))
	text(c(3), cex=cex_plot, y_lim+3*y_offset, c("QualSyst"))
	text(c(3), cex=cex_plot, y_lim-y_offset, c("score"))
	text(c(3.5), cex=cex_plot, y_lim-y_offset, c("weight"))
		
	### Switch to bold italic font
	par(font=4)

	### add text for subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.1, c(
								   sprintf("Selegiline (n = %s)", sum(dat_maoRCT$ni[dat_maoRCT$treatment_diff=="selegiline"])), 
								   sprintf("Safinamide (n = %s)", sum(dat_maoRCT$ni[dat_maoRCT$treatment_diff=="safinamide"])),
								   sprintf("Rasagiline (n = %s)", sum(dat_maoRCT$ni[dat_maoRCT$treatment_diff=="rasagiline"]))))

	### set par back to the original settings
	par(op)
	
	### fit random-effects model in all subgroups
	res.saf <- rma(yi, vi, subset=(treatment_diff=="safinamide"), data=dat_maoRCT)
	res.ras <- rma(yi, vi, subset=(treatment_diff=="rasagiline"), data=dat_maoRCT)
	res.sel <- rma(yi, vi, subset=(treatment_diff=="selegiline"), data=dat_maoRCT)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.ras, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.saf), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.saf, row= 12.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ras), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.sel, row= 17.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.sel), cex=cex_plot*fac_cex, efac = c(.5))
	
	abline(h=0)
	
	### fit meta-regression model to test for subgroup differences
	dat_maoRCT$treatment_diff <- as.factor(dat_maoRCT$treatment_diff)
	groups 	<- c("selegiline", "safinamide", "rasagiline")
	levels(dat_maoRCT) <- groups
	res_overall_maoRCT <- rma(yi, vi, mods = ~ treatment_diff - 1, data=dat_maoRCT)
	
	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overall_maoRCT$QM, digits=2, format="f")), ", df = ", .(res_overall_maoRCT$p - 1),
		 ", p = ", .(formatC(res_overall_maoRCT$QMp, digits=3, format="f")))))
	dev.off()

	# Run post-hoc analyses
	# =================	
	pwc 					<- summary(glht(res_overall_maoRCT, linfct=cbind(contrMat(rep(1,3), type="Tukey"))), 
							test=adjusted("BH")) # pairwise comparisons
	mat2plot 				<- matrix(0, nrow = 3, ncol = 3, dimnames=list(toupper(substr(groups, 1, 3)), toupper(substr(groups, 1, 3)))) + + diag(3)/1000000
	mat_significance 		<- mat2plot 
	mat2plot[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$tstat)
	mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
	melted_cormat 			<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
	melted_cormat$pval 		<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
	melted_cormat$pval 		<- melted_cormat$pval$value # dodgy solution
	colnames(melted_cormat)	<-c("Var1", "Var2", "tstat", "pval")
	melted_cormat[melted_cormat==0] <- NA # remove zero values and convert to NA
	melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001] <- "*"
	melted_cormat$pval[melted_cormat$pval<.001] <- "**"
	melted_cormat$pval[melted_cormat$pval>.05] <- NA

	# Create a heatmap for pairwise comparisons
	svg(filename=file.path(wdir, "results", "heatmaps.maoRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	ggplot(data = melted_cormat %>% drop_na(-pval), aes(Var2, Var1, fill = tstat))+
		geom_tile(color = "white")+
		scale_fill_gradient2(mid="#FBFEF9",low="#0C6291",high="#A63446", limits=c(-5,5),space = "Lab", 
						   name="T-statistics") +
		geom_text(aes(label = pval)) +
		ylab(NULL)  + xlab(NULL) +
		ggtitle("Pairwise comparison between MAO-inhibitors") + 
		theme_minimal() +
		#theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   #size = 12, hjust = 1), 
									   #axis.text.x=element_blank()) +
		coord_fixed()
	dev.off()

	
	# ==================================================================================================
	# ==================================================================================================
	# E. Three-level hierarchical model with all studies included for adverse events
	# ==================================================================================================
	# ==================================================================================================
	
	# Load data
	dat_adverse_events 			<- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
	dat_adverse_events 			<- dat_adverse_events %>% na_if("NA") %>% drop_na(ae1i, ae2i, n1i, n2i) %>% 
									mutate_at(c('ae1i', 'ae2i','do1i', 'do2i', 'n1i', 'n2i', 
										'age1', 'age2', 'updrs1', 'updrs2'), as.numeric)# drop NA values in ae1i convert to numeric
	dat_adverse_events 			<- dat_adverse_events %>% dplyr::filter(.,study_type=="RCT")
	dat_adverse_events$study.id <- 1:dim(dat_adverse_events)[1]
	attr(dat_adverse_events$yi, 'slab')<- dat_adverse_events$slab			
	
	dat_adverse_events$order_no[dat_adverse_events$order_no==5] = 1
	dat_adverse_events$order_no[dat_adverse_events$order_no==6] = 2
	dat_adverse_events$order_no[dat_adverse_events$order_no==7] = 3
	dat_adverse_events$order_no[dat_adverse_events$order_no==8] = 4
	dat_adverse_events$order_no[dat_adverse_events$order_no==10] = 5
	dat_adverse_events$order_no[dat_adverse_events$order_no==12] = 6

	dat_adverse_events <- dat_adverse_events %>% drop_na(do1i, do2i)
	dat_adverse_events$study.id <- 1:dim(dat_adverse_events)[1]

	
	
	sorted_indices 	<- dat_adverse_events %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 
	
	# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
	# =================
	xrows_temp 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()

	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	dat_adverse_events 			<- escalc(measure="PETO", ai=dat_adverse_events$do1i, 
														bi=dat_adverse_events$n1i - dat_adverse_events$do1i, 
														ci=dat_adverse_events$do2i, 
														di=dat_adverse_events$n2i - dat_adverse_events$do2i,
														n1i=n1i, n2i=n2i,
														add=1/2, to="only0", data=dat_adverse_events)

	# Impute variance-covariance-matrix according to correlated effects (see above)
	# =================
	V_matAE 	<- impute_covariance_matrix(vi = dat_adverse_events$vi, cluster=dat_adverse_events$study, r = .7)

	fit 		<- rma.mv(yi, slab=dat_adverse_events$slab, V=V_matAE, random = ~ 1 | as.factor(category)/study.id,
							data=dat_adverse_events, mods= ~ qualsyst, method="REML")
	
	cex_plot 	<- 1 		# plot size
	y_lim 		<- 50 	# use last "start" as y-limit for forest plot
	y_offset 	<- .5		# offset used to plot the header of the column
	# svg(filename=file.path(wdir, "results", "meta_analysis.agonistsRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	forest(fit, xlim=c(-5, 10), at=log(c(0.25, 1, 5)), atransf=exp,
			ilab=cbind(dat_adverse_events$ni, dat_adverse_events$treatment, dat_adverse_events$study_type, 
			sprintf("%.02f",  dat_adverse_events$qualsyst), 
			paste0(formatC(weights(fit), format="f", digits=1, width=4), "%")),
			xlab="Odds ratio",
			ilab.xpos=c(-3, -2, 2.5, 3, 3.5), 
			cex=cex_plot, ylim=c(-1, y_lim),
			rows=xrows,
			efac =c(.8), #offset=y_offset, 
			mlab=mlabfun("RE Model for All Studies", res_agonistsRCT),
			font=4, header="Author(s) and Year")
	
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(font=2, mar = c(2, 2, 2, 2))

	text((c(-.5, .5)), cex=cex_plot, y_lim, c("Favours placebo", "Favors intervention"), 
			pos=c(2, 4), offset=y_offset)

	### Add column headings for everything defined in "forest::ilab"
	text(c(-3), cex=cex_plot, y_lim, c("n"))
	text(c(2.5), cex=cex_plot, y_lim, c("type"))
	text(c(3), cex=cex_plot, y_lim+2*y_offset, c("QualSyst"))
	text(c(3), cex=cex_plot, y_lim, c("score"))
	text(c(3.5), cex=cex_plot, y_lim, c("weight"))	
	
	### add text for subgroups
	### add text for the subgroups
	text(-5, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.5, c(
								   sprintf("Dopamine agonists (n = %s)", sum(dat_adverse_events$ni[dat_adverse_events$category=="dopamine_agonists"])),
								   sprintf("Clozapine (n = %s)", sum(dat_adverse_events$ni[dat_adverse_events$category=="clozapine"])), 
								   sprintf("MAO-inhibitor (n = %s)", sum(dat_adverse_events$ni[dat_adverse_events$category=="mao_inhibitors"])), 
								   sprintf("Budipine (n = %s)", sum(dat_adverse_events$ni[dat_adverse_events$category=="budipine"])),				
								   sprintf("Adenosine antagonists (n = %s)", sum(dat_adverse_events$ni[dat_adverse_events$category=="adenosine_antagonist"]))
									))

	### set par back to the original settings
	par(op)
	
	### fit random-effects model in all subgroups
	res.dop <- rma(yi, vi, subset=(category=="dopamine_agonists"), data=dat_adverse_events)
	res.clo <- rma(yi, vi, subset=(category=="clozapine"), data=dat_adverse_events)
	res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"), data=dat_adverse_events)
	res.bud <- rma(yi, vi, subset=(category=="budipine"), data=dat_adverse_events)
	res.ade <- rma(yi, vi, subset=(category=="adenosine_antagonist"), data=dat_adverse_events)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.ade, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ade), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.bud, row= 6.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bud), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.mao, row= 11.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.clo, row= 25.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.dop, row= 30.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.dop), cex=cex_plot*fac_cex, efac = c(.5))
	
	abline(h=0)

	### fit meta-regression model to test for subgroup differences
	dat_adverse_events$category <- as.factor(dat_adverse_events$category)
	groups 	<- c("adenosine_antagonist", "budipine", "mao_inhibitors", "clozapine", "dopamine_agonists")
	levels(dat_adverse_events) <- groups
	res_overallAE <- rma(yi, vi, mods = ~ category - 1, data=dat_adverse_events)
	
	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overallAE$QM, digits=2, format="f")), ", df = ", .(res_overallAE$p - 1),
		 ", p = ", .(formatC(res_overallAE$QMp, digits=3, format="f")))))
	dev.off()

	# Run post-hoc analyses
	# =================	
	pwc 					<- summary(glht(res_overallAE, linfct=cbind(contrMat(rep(1,5), type="Tukey"))), 
							test=adjusted("BH")) # pairwise comparisons
	mat2plot 				<- matrix(0, nrow = 5, ncol = 5, dimnames=list(toupper(substr(groups, 1, 5)), toupper(substr(groups, 1, 5)))) + + diag(5)/1000000
	mat_significance 		<- mat2plot 
	mat2plot[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$tstat)
	mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
	melted_cormat 			<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
	melted_cormat$pval 		<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
	melted_cormat$pval 		<- melted_cormat$pval$value # dodgy solution
	colnames(melted_cormat)	<-c("Var1", "Var2", "tstat", "pval")
	melted_cormat[melted_cormat==0] <- NA # remove zero values and convert to NA
	melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001] <- "*"
	melted_cormat$pval[melted_cormat$pval<.001] <- "**"
	melted_cormat$pval[melted_cormat$pval>.05] <- NA

	# Create a heatmap for pairwise comparisons
	# svg(filename=file.path(wdir, "results", "heatmaps.maoRCT.v1.0.svg"))
	windowsFonts("Arial" = windowsFont("TT Arial"))
	ggplot(data = melted_cormat %>% drop_na(-pval), aes(Var2, Var1, fill = tstat))+
		geom_tile(color = "white")+
		scale_fill_gradient2(mid="#FBFEF9",low="#0C6291",high="#A63446", limits=c(-5,5),space = "Lab", 
						   name="T-statistics") +
		geom_text(aes(label = pval)) +
		ylab(NULL)  + xlab(NULL) +
		ggtitle("Pairwise comparison between MAO-inhibitors") + 
		theme_minimal() +
		#theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   #size = 12, hjust = 1), 
									   #axis.text.x=element_blank()) +
		coord_fixed()
	dev.off()

	
	
	# ==================================================================================================
	# ==================================================================================================
	# F. Summary of data from all studies
	# ==================================================================================================
	# ==================================================================================================	
	
	dat_table_one 				<- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
	dat_table_one 				<- dat_table_one %>% mutate_at(c('age1', 'age2', 'n1i', 'n2i', 'updrs1', 'updrs2', 'qualsyst'), as.numeric) %>% 
									mutate(age_collated = (n1i*age1 + n2i*age2)/(n1i + n2i))
	dat_table_one$age_collated 	<- ifelse(is.na(dat_table_one$age_collated), dat_table_one$age1, dat_table_one$age_collated)

	dat_table_one 				<- dat_table_one %>% mutate(updrs_collated = (n1i*updrs1 + n2i*updrs2)/(n1i + n2i))
	dat_table_one$updrs_collated<- ifelse(is.na(dat_table_one$updrs_collated), dat_table_one$updrs1, dat_table_one$updrs_collated)
	dat_table_one$study_type 	<- as.factor(dat_table_one$study_type)
	dat_table_one$category 	<- as.factor(dat_table_one$category)
	
	NumVars <- c("age_collated", "qualsyst", "updrs_collated", "ni")
	catVars <- c("study_type", "category")
	myVars 	<- c(catVars, NumVars)
	tableOne <- CreateTableOne(vars = myVars, data = dat_table_one, factorVars = catVars)
	write.csv(print(tableOne), file=file.path(getwd(), "results", "tableOne.v1.0.csv"))
	
	
	# ANOVA for UPDRS
	dat_adverse_events 	<- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
	dat_anova <- dat_adverse_events %>% dplyr::select(slab:study_type, age1, age2, n1i, n2i, updrs1, updrs2) %>% na_if("NA") %>% drop_na(-n2i, -updrs2, -age2)
	
	groups 		<- data.frame(	levodopa=c("levodopa", rep(NA, 8)), # adds groups to dataframe
								betablocker=c("propanolol", "zuranolone", "nipradilol", rep(NA, 6)),
								primidone=c("primidone", rep(NA, 8)),
								anticholinergics=c("trihexphenidyl", "benztropin", rep(NA,7 )),
								amantadine=c("amantadine", rep(NA, 8)), 
								dopamine_agonists=c("ropinirole", "rotigotine", "apomorphine", "pramipexole", 
													"piribedil", "bromocriptine", "pergolide", "cabergoline", 
													"dihydroergocriptin"), 
								clozapine=c("clozapine", rep(NA, 8)), 
								mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)),
								cannabis=c("cannabis", rep(NA, 8)), 
								budipine=c("budipine", rep(NA, 8)),
								zonisamide=c("zonisamide", rep(NA, 8)),
								adenosine_antagonist=c("theophylline", "caffeine", "adenosineA2a", rep(NA, 6)),
								botox=c("botulinum", rep(NA, 8)))

	dat_anova 	<- dat_anova %>% mutate(category=NA, age=NA, updrs=NA, class=NA, order_no=NA) # add columns to dataframe to work with
	iter		<- 0
	for (k in colnames(groups)){ # assigns category to the different treatments
		iter 					<- iter + 1
		temp 					<- dat_anova %>% rowid_to_column %>% 
									filter(if_any(everything(),~str_detect(., c(paste(groups[[k]][!is.na(groups[[k]])],collapse='|'))))) 
		dat_anova$category[temp$rowid]<- colnames(groups)[iter]
		dat_anova$order_no[temp$rowid]<- iter
	}
	
	
	# ==================================================================================================
	# ==================================================================================================
	# G. Anova for suvgroups of different agents
	# ==================================================================================================
	# ==================================================================================================	

	# Load data and extract further information needed (e.g., total age, cumulative updrs between groups, etc.)
	# =================	
	dat_anova 		<- dat_anova %>% mutate_at(c('age1', 'age2', 'n1i', 'n2i', 'updrs1', 'updrs2'), as.numeric) %>% 
									mutate(age = (n1i*age1 + n2i*age2)/(n1i + n2i))
	dat_anova$age	<- ifelse(is.na(dat_anova$age), dat_anova$age1, dat_anova$age)
	dat_anova 		<- dat_anova %>% mutate(updrs = (n1i*updrs1 + n2i*updrs2)/(n1i + n2i))
	dat_anova$updrs	<- ifelse(is.na(dat_anova$updrs), dat_anova$updrs1, dat_anova$updrs)

	# Run ANOVA for age and for UPDRS
	# =================		
	stat.test.age 		<- dat_anova %>% anova_test(age ~ category)
	post_hoc.age 		<- aov(age ~ category, data=dat_anova) %>% tukey_hsd()

	stat.test.updrs 	<- dat_anova %>% anova_test(updrs ~ category)
	post_hoc.updrs 		<- aov(updrs ~ category, data=dat_anova) %>% tukey_hsd()
	significant_pairs 	<- post_hoc.updrs[post_hoc.updrs$p.adj.signif!="ns",]

	svg(filename=file.path(wdir, "results", "group_comparisons.v1.0.svg"))
	a 	<- ggboxplot(dat_anova, x = "category", y = "age") +
		stat_compare_means(method = "anova", label.y = 75)+ 
		xlab("") + ylim(50, 75) +
		  theme(axis.ticks.x = element_blank(),
        axis.text.x = element_blank()) + 
		geom_hline(yintercept = mean(dat_anova$age), linetype = 2)
	  
	b <- ggboxplot(dat_anova, x = "category", y = "updrs") +
		xlab("") + ylim(0, 60) +
		rotate_x_text(angle = 45) +
		stat_compare_means(method = "anova", label.y = 60) + 
		geom_hline(yintercept = mean(dat_anova$updrs), linetype = 2)
	
	a/b # plot one below the other
	dev.off()