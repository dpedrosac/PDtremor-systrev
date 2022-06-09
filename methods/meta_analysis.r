	# This is code to run meta analyses on the results from the systematic literature review
	# For further details on the data extraction cf. worksheets in __included.xlsx and 
	# code in analysis_dataframe.r 
	# Code developed by David Pedrosa

	# Version 1.1 # 2022-04-29, minor changes in the layout

	# ==================================================================================================
	## In case of multiple people working on one project, this helps to create an automatic script
	username = Sys.info()["login"]
	if (username == "dpedr") {
	wdir = "D:/Jottacloud/PDtremor_syst-review/"
	} else if (username == "david") {
	wdir = "/media/storage/Jottacloud/PDtremor_syst-review/"
	}
	setwd(wdir)

	# ==================================================================================================
	## Specify packages of interest and load them automatically if needed
	packages = c("readxl", "dplyr", "plyr", "tibble", "tidyr", "stringr", "openxlsx", 
					"metafor", "tidyverse", "clubSandwich", "multcomp", "reshape2", "latex2exp") 											# packages needed

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


	# ==================================================================================================
	# ==================================================================================================
	# A. Three-level hierarchical model including all studies for effect sizes and category as 2nd level
	# ==================================================================================================
	# ==================================================================================================

	# Load data
	dat_results <- read.csv(file.path(wdir, "results", "results_meta-analysis.csv"), 
								stringsAsFactors = F,  encoding="UTF-8")
	cex_plot 	<- 1
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

	# Prepare some additional data for plotting later
	dat <- dat_results %>% mutate(category=NA, class=NA, order_no=NA) # add columns to dataframe to work with
	iter<- 0
	for (k in colnames(groups)){ # assigns category to the different treatments
		iter 					<- iter + 1
		comparator_groups 		<- groups[[k]][!is.na(groups[[k]])]
		temp 					<- dat_results %>% rowid_to_column %>% 
								filter(if_any(everything(),~str_detect(., c(paste(comparator_groups,collapse='|'))))) 
		dat$category[temp$rowid]<- colnames(groups)[iter]
		dat$order_no[temp$rowid]<- iter
	}

	sorted_indices 	<- dat %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 
	xrows_all 		<- 1:150
	spacing 		<- 5 
	start 			<- 3
	xrows 			<- c()
	headings 		<- c()
	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_all[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	
	y_lim = start
	dat 				<- dat %>% drop_na(-class)
	dat$study 			<- 1:dim(dat)[1] # add the study ID in the daat frame
	dat 				<- dat %>% arrange(desc(order_no), treatment)
	attr(dat$yi, 'slab')<- dat$slab			
	temp 				<- dat %>% rowid_to_column %>% 
								filter(if_any(everything(),
										~str_detect(., c(paste(c("levodopa", "amantadine", "clozapine", 
																	"budipine", "cannabis", "zonisamide", 
																	"primidone", "botox"), collapse='|'))))) 
	dat$treatment2 <- rep(NA, dim(dat)[1])
	dat$treatment2[setdiff(1:dim(dat)[1], temp$rowid)] = dat$treatment[setdiff(1:dim(dat)[1], temp$rowid)]

	coal = TRUE
	if (coal==TRUE) {
		dat <- dat %>% mutate(treatment2= coalesce(treatment2, treatment))
	}

	covariance_adaptation = TRUE # in case of studies with the same participants, covariance should be adopted
	if (covariance_adaptation==TRUE) {
		dat$study[which(dat$slab=="Koller et al. 1987.1")] <- dat$study[which(dat$slab=="Koller et al. 1987")]
		dat$study[which(dat$slab=="Brannan & Yahr (1995)")] <- dat$study[which(dat$slab=="Brannan & Yahr (1995)")][1]
		dat$study[which(dat$slab=="Sahoo et al. 2020")] <- dat$study[which(dat$slab=="Sahoo et al. 2020")][1]
		dat$study[which(dat$slab=="Navan et al. 2003.1")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][1]
		dat$study[which(dat$slab=="Navan et al. 2003")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][2]
		dat$study[which(dat$slab=="Friedman et al. 1997")] <- dat$study[which(dat$slab=="Friedman et al. 1997")][1]
		dat$study[which(dat$slab=="Zach et al. 2020")] <- dat$study[which(dat$slab=="Dirkx et al. 2019")][1]
	}

	# Input the Variance Covariance matrix according to correlated effects (see above
	V_mat <- impute_covariance_matrix(vi = dat$vi, cluster=dat$study, r = .7)

	res <- rma.mv(yi, slab=dat$slab, V=V_mat,
					random = ~1 | category/study, #list(~ 1 | factor(category), ~ 1 | factor(study)) , #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst*study_type, tdist=TRUE, #struct="DIAG",
					data=dat, verbose=TRUE, control=list(optimizer="optim", optmethod="Nelder-Mead"))

	# Run some diagnostics on the model with all studies/categories
	par(mfrow=c(2,1))
	profile(res, sigma2=1, steps=50)
	profile(res, sigma2=2, steps=50)
	dev.copy(svg,file.path(wdir, "results", "suppl_figure.diagResAll.svg"),width = 8.27,height = 11.69)
	# dev.off()

	# Calculation of I^2 statistic
	W <- diag(1/res$vi)
	X <- model.matrix(res)
	P <- W - W %*% X %*% solve(t(X) %*% W %*% X) %*% t(X) %*% W
	100 * sum(res$sigma2) + (sum(res$sigma2) + (res$k-res$p)/sum(diag(P)))
	100 * res$sigma2 / (sum(res$sigma2) + (res$k-res$p)/sum(diag(P)))

	# Start plotting data using a forest plot
	forest(res, xlim=c(-10, 4.6), at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
		   ilab=cbind(sprintf("%.02f",  dat$qualsyst), dat$ni, 
		   paste0(formatC(weights(res), format="f", digits=1, width=4), "%"),
		   dat$treatment2, dat$study_type), #, dat$tneg, dat$cpos, dat$cneg),
		   xlab="Standardized mean change", # TODO: Latrex code for exp SMCR
		   ilab.xpos=c(2.5, -5, 3, -4, 2), 
		   #ilab.xpos=c(-9.5,-8,-6,-4.5), 
		   cex=cex_plot, ylim=c(-1, y_lim),
		   rows=xrows, #c(3:4, 9:18, 23:27, 32:64, 69, 74:81),
		   #slab=dat$slab, 
		   #order=dat$order_no,
		   efac =c(.75),
		   mlab=mlabfun("RE Model for All Studies", res),
		   font=4, header="Author(s) and Year")
	 
	### set font expansion factor (as in forest() above) and use a bold font
	op <- par(cex=0.5, font=4, mar = c(2, 2, 2, 2))

	text(log(c(.1, 3.5)), cex=2, y_lim, c("Tremor suppression","Tremor increase"), pos=c(4,2), offset=-0.5)

	### add additional column headings to the plot
	text(c(2.5), cex=2, y_lim, c("QualSyst"))
	text(c(2.5), cex=2, y_lim-1, c("score"))
	text(c(3), cex=2, y_lim-1, c("weight"))
	text(c(-5), cex=2, y_lim-1, c("n"))
	text(c(-4), cex=2, y_lim-1, c("agent"))
	text(c(2), cex=2, y_lim-1, c("type"))
	 
	### switch to bold italic font
	par(font=4)
	 
	### add text for the subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=2, c(sprintf("Levodopa (n = %s)", sum(dat$ni[dat$category=="levodopa"])),
								   sprintf("beta-blocker (n = %s)", sum(dat$ni[dat$category=="betablocker"])),
								   sprintf("Primidone (n = %s)", sum(dat$ni[dat$category=="primidone"])),
								   sprintf("Anticholinergics (n = %s)", sum(dat$ni[dat$category=="anticholinergics"])),
								   sprintf("Amantadine (n = %s)", sum(dat$ni[dat$category=="amantadine"])),
								   sprintf("Dopamine agonists (n = %s)", sum(dat$ni[dat$category=="dopamine_agonists"])), 
								   sprintf("Clozapine (n = %s)", sum(dat$ni[dat$category=="clozapine"])), 
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
	res.clo <- rma(yi, vi, subset=(category=="clozapine"),  data=dat)
	res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"),     data=dat)
	res.cbd <- rma(yi, vi, subset=(category=="cannabis"),  data=dat)
	res.bud <- rma(yi, vi, subset=(category=="budipine"),  data=dat)
	res.zon <- rma(yi, vi, subset=(category=="zonisamide"),  data=dat)
	res.ade <- rma(yi, vi, subset=(category=="adenosine_antagonist"),  data=dat)
	res.bot <- rma(yi, vi, subset=(category=="botox"),  data=dat)

	par(font=4)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.lev, row= 118.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.bet, row= 111.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pri, row= 106.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pri), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.cho, row= 100.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.ama, row= 94.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ama), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.clo, row= 59.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.dag, row= 50.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.dag), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.mao, row= 35.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.cbd, row= 30.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.cbd), cex=cex_plot*.fac_cex, efac = c(.5))
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
	#addpoly(res_sg, row= -1.8, mlab=mlabfun("RE Model for Subgroup", res_sg), cex=cex_plot*fac_cex, efac = c(.5))

	funnel_contour <- funnel(res, level=c(90, 95, 99), back="gray95",
                  shade=c("white", "gray75", "gray85", "gray95"), 
                  refline=0, legend="topleft")
	
	dat$residuals	<- residuals.rma(res)
	dat$precision	<- 1/sqrt(dat$vi)
	egger_res		<- rma.mv(residuals~precision,vi,data=dat, random = ~ factor(study) | factor(category))
	
	par(mfrow=c(2,1))
	profile(res, tau2=1)
	profile(res, rho=1, xlim=c(0.01,0.9))
	
	# ==================================================================================================
	# ==================================================================================================
	# B. Two-level hierarchical model with studies using dopamine agonists
	# ==================================================================================================
	# ==================================================================================================

	# Settings
	y_lim  		<- 64
	headings	<- c(8.5, 18.5, 30.5, 38.5, 45.5, 60 ) 
	cex_plot  	<- 1

	dat_agonists <- dat %>% filter(category %in% c("levodopa", "dopamine_agonists"))	
	remove_ergot_derivates=TRUE
	if (remove_ergot_derivates){
		dat_agonists <- dat_agonists %>% filter(treatment2!="dihydroergocriptin") %>% 
			filter(treatment2!="cabergoline") %>% 
			filter(treatment2!="pergolide")
	}
	
	# Imput the Variance Covariance matrix according to correlated effects (see above
	V_mat_agonists <- impute_covariance_matrix(vi = dat_agonists$vi, cluster=dat_agonists$study, r = .7)

	res_agonists <- rma.mv(yi, slab=dat_agonists$slab, V=V_mat_agonists,
					random = ~ factor(study) | factor(treatment), #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst*study_type, tdist=TRUE, #struct="DIAG",
					data=dat_agonists, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
	
	# Start plotting data using a forest plot
	forest(res_agonists, xlim=c(-10, 4.6), at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
		   ilab=cbind(sprintf("%.02f",  dat_agonists$qualsyst), dat_agonists$ni, 
		   paste0(formatC(weights(res_agonists), format="f", digits=1, width=4), "%"),
		   dat_agonists$treatment2, dat_agonists$study_type), #, dat$tneg, dat$cpos, dat$cneg),
		   ilab.xpos=c(2.5, -5, 3, -4, 2), 
		   xlab="Standardized mean change", #ilab.xpos=c(-9.5,-8,-6,-4.5), 
		   cex=cex_plot, ylim=c(-1, y_lim),
		   rows=c(3:6, 11:16, 21:28, 33:36, 41:43, 48:57),
		   #slab=dat$slab, 
		   #order=dat$order,
		   efac =c(.75),
		   mlab=mlabfun("RE Model for All Studies", res_agonists),
		   font=4, header="Author(s) and Year")

	### set font expansion factor (as in forest() above) and use a bold font
	#op <- par(cex=0.5, font=4, mar = c(2, 2, 2, 2))
	
	### add additional column headings to the plot
	text(c(2.5), cex=2, y_lim, c("QualSyst"))
	text(c(2.5), cex=2, y_lim-1, c("score"))
	text(c(3), cex=2, y_lim-1, c("weight"))
	text(c(-5), cex=2, y_lim-1, c("n"))
	text(c(-4), cex=2, y_lim-1, c("agent"))
	text(c(2), cex=2, y_lim-1, c("type"))
		
	### fit random-effects model in all subgroups
	res.apo <- rma(yi, vi, subset=(treatment2=="apomorphine"), data=dat_agonists)
	# res.cab <- rma(yi, vi, subset=(treatment2=="cabergoline"), data=dat_agonists)
	# res.per <- rma(yi, vi, subset=(treatment2=="pergolide"), data=dat_agonists)
	#res.dih <- rma(yi, vi, subset=(treatment2=="dihydroergocriptin"), data=dat_agonists)
	res.rot <- rma(yi, vi, subset=(treatment2=="rotigotine"), data=dat_agonists)
	res.rop <- rma(yi, vi, subset=(treatment2=="ropinirole"), data=dat_agonists)
	res.pir <- rma(yi, vi, subset=(treatment2=="piribedil"), data=dat_agonists)
	res.pra <- rma(yi, vi, subset=(treatment2=="pramipexole"), data=dat_agonists)
	res.lev <- rma(yi, vi, subset=(treatment2=="levodopa"), data=dat_agonists)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.apo, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.apo), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pir, row= 9.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pir), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pra, row= 19.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pra), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.rop, row= 31.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rop), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.rot, row= 39.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rot), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.lev, row= 46.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*fac_cex, efac = c(.5))
	
	### add text for the subgroups
	text(-10, sort(headings, decreasing=TRUE), pos=4, cex=2, c(sprintf("Levodopa (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="levodopa"])),
								   sprintf("Rotigotine (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="rotigotine"])),
								   sprintf("Ropinirole (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="ropinirole"])),
								   sprintf("Pramipexole (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="pramipexole"])),
								   sprintf("Piribedil (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="piribedil"])),
								   sprintf("Apomorphine (n = %s)", sum(dat_agonists$ni[dat_agonists$treatment2=="apomorphine"]))))

	### fit meta-regression model to test for subgroup differences
	dat_agonists$treatment2 <- as.factor(dat_agonists$treatment2)
	groups <- c("levodopa", "apomorphine", "piribedil", "pramipexole", "ropinirole", "rotigotine")
	levels(dat_agonists) <- groups
	res_overall_agonists <- rma(yi, vi, mods = ~ treatment2 - 1, data=dat_agonists)

	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overall_agonists$QM, digits=2, format="f")), ", df = ", .(res_overall_agonists$p - 1),
		 ", p = ", .(formatC(res_overall_agonists$QMp, digits=3, format="f")))))
	
	# Run posthoc analyses and plot results as heatmap 
	pwc 			<- summary(glht(res_overall_agonists, linfct=cbind(contrMat(rep(1,6), type="Tukey"))), test=adjusted("BH")) # pairwise comparisons
	mat <- matrix(0, nrow = 6, ncol = 6, dimnames=list(groups, groups))
	mat[lower.tri(mat, diag = FALSE)] <- unname(pwc$test$pvalues)
	
	melted_cormat <- melt(mat, na.rm = TRUE)

	# Heatmap
	ggplot(data = melted_cormat, aes(Var2, Var1, fill = value))+
	  geom_tile(color = "white")+
	  scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
						   midpoint = 0, limit = c(0,1), space = "Lab", 
						   name="Significance \n (p_val)") +
	  geom_text(aes(label = round(value, 3))) +
	  theme_minimal()+ 
	  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   size = 12, hjust = 1)) +
  coord_fixed()
	
	
	# ==================================================================================================
	# ==================================================================================================
	# C. Two-level hierarchical model with studies using dopamine agonists (only RCTs)
	# ==================================================================================================
	# ==================================================================================================

	dat_agonistsRCT <- dat %>% filter(category %in% c("levodopa", "dopamine_agonists")) %>% filter(study_type=="RCT")
	remove_ergot_derivates=TRUE
	if (remove_ergot_derivates){
		dat_agonistsRCT <- dat_agonistsRCT %>% filter(treatment2!="dihydroergocriptin") %>% 
			filter(treatment2!="cabergoline") %>% 
			filter(treatment2!="pergolide")
	}
	
	# Input the Variance Covariance matrix according to correlated effects (see above
	V_mat_agonists 	<- impute_covariance_matrix(vi = dat_agonistsRCT$vi, cluster=dat_agonistsRCT$study, r = .7)

	res_agonistsRCT	<- rma.mv(yi, slab=dat_agonistsRCT$slab, V=V_mat_agonists,
					random = ~ factor(study) | factor(treatment), #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG",
					data=dat_agonistsRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))

	dat_agonistsRCT$order_no 	<- as.integer(as.factor(sort(dat_agonistsRCT$treatment)))
	sorted_indices 				<- dat_agonistsRCT %>% dplyr::select(order_no) %>% 
								drop_na() %>% gather(order_no) %>% count() %>% arrange(desc(order_no))# 

	xrows_all 			<- 1:150
	spacing 			<- 5 
	start 				<- 3
	xrows 				<- c()
	headings 			<- c()
	for (i in 1:dim(sorted_indices)[1]){
		vector_group 	<- xrows_all[start:(start+sorted_indices[i,2]-1)]
		xrows 			<- c(xrows, vector_group)
		headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
		start 			<- start+sorted_indices[i,2] + 4
	} 
	y_lim = start
	dat_agonistsRCT <- dat_agonistsRCT %>% arrange(desc(order_no), treatment)
	
	# Settings for plotting
	cex_plot  	<- 1
	
	# Start plotting data using a forest plot
	forest(res_agonistsRCT, xlim=c(-10, 4.6), at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
		   ilab=cbind(sprintf("%.02f",  dat_agonistsRCT$qualsyst), dat_agonistsRCT$ni, 
		   paste0(formatC(weights(res_agonistsRCT), format="f", digits=1, width=4), "%"),
		   dat_agonistsRCT$treatment2, dat_agonistsRCT$study_type), #, dat$tneg, dat$cpos, dat$cneg),
		   ilab.xpos=c(2.5, -5, 3, -4, 2), 
		   #ilab.xpos=c(-9.5,-8,-6,-4.5), 
		   cex=cex_plot, ylim=c(-1, y_lim),
		   rows=xrows,
		   #slab=dat$slab, 
		   #order=dat$order,
		   efac =c(.75),
		   mlab=mlabfun("RE Model for All Studies", res_agonistsRCT),
		   font=4, header="Author(s) and Year")

	### set font expansion factor (as in forest() above) and use a bold font
	#op <- par(cex=0.5, font=4, mar = c(2, 2, 2, 2))
	
	### add additional column headings to the plot
	text(c(2.5), cex=2, y_lim, c("QualSyst"))
	text(c(2.5), cex=2, y_lim-1, c("score"))
	text(c(3), cex=2, y_lim-1, c("weight"))
	text(c(-5), cex=2, y_lim-1, c("n"))
	text(c(-4), cex=2, y_lim-1, c("agent"))
	text(c(2), cex=2, y_lim-1, c("type"))
	
	### fit random-effects model in all subgroups
	res.apo <- rma(yi, vi, subset=(treatment2=="apomorphine"), data=dat_agonistsRCT)
	# res.cab <- rma(yi, vi, subset=(treatment2=="cabergoline"), data=dat_agonists)
	# res.per <- rma(yi, vi, subset=(treatment2=="pergolide"), data=dat_agonists)
	#res.dih <- rma(yi, vi, subset=(treatment2=="dihydroergocriptin"), data=dat_agonists)
	res.rot <- rma(yi, vi, subset=(treatment2=="rotigotine"), data=dat_agonistsRCT)
	res.rop <- rma(yi, vi, subset=(treatment2=="ropinirole"), data=dat_agonistsRCT)
	res.pir <- rma(yi, vi, subset=(treatment2=="piribedil"), data=dat_agonistsRCT)
	res.pra <- rma(yi, vi, subset=(treatment2=="pramipexole"), data=dat_agonistsRCT)
	#res.lev <- rma(yi, vi, subset=(treatment2=="levodopa"), data=dat_agonists)

	yshift = .2
	fac_cex = 1
	### add summary polygons for the three subgroups
	addpoly(res.rot, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rot), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.rop, row= 7.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.rop), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pra, row= 15.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pra), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.pir, row= 24.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.pir), cex=cex_plot*fac_cex, efac = c(.5))
	addpoly(res.apo, row= 30.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.apo), cex=cex_plot*fac_cex, efac = c(.5))
	
	### add text for the subgroups
	text(-10, sort(headings, decreasing=F), pos=4, cex=2, c(sprintf("Rotigotine (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment2=="rotigotine"])),
								   sprintf("Ropinirole (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment2=="ropinirole"])),
								   sprintf("Pramipexole (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment2=="pramipexole"])),
								   sprintf("Piribedil (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment2=="piribedil"])),
								   sprintf("Apomorphine (n = %s)", sum(dat_agonistsRCT$ni[dat_agonistsRCT$treatment2=="apomorphine"]))))

	### fit meta-regression model to test for subgroup differences
	dat_agonistsRCT$treatment2 <- as.factor(dat_agonistsRCT$treatment2)
	groups <- c("apomorphine", "piribedil", "pramipexole", "ropinirole", "rotigotine")
	levels(dat_agonistsRCT) <- groups
	res_overall_agonistsRCT <- rma(yi, vi, mods = ~ treatment2 - 1, data=dat_agonistsRCT)
	
	### add text for the test of subgroup differences
	text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
		 "Q[M]", " = ", .(formatC(res_overall_agonistsRCT$QM, digits=2, format="f")), ", df = ", .(res_overall_agonistsRCT$p - 1),
		 ", p = ", .(formatC(res_overall_agonistsRCT$QMp, digits=3, format="f")))))
	
	# Run posthoc analyses and plot results as heatmap 
	pwc 	<- summary(glht(res_overall_agonistsRCT, linfct=cbind(contrMat(rep(1,5), type="Tukey"))), test=adjusted("BH")) # pairwise comparisons
	mat 	<- matrix(0, nrow = 5, ncol = 5, dimnames=list(groups, groups))
	mat[lower.tri(mat, diag = FALSE)] <- unname(pwc$test$pvalues)
	
	melted_cormat <- melt(mat, na.rm = TRUE)

	# Heatmap
	ggplot(data = melted_cormat, aes(Var2, Var1, fill = value))+
	  geom_tile(color = "white")+
	  scale_fill_gradient2(low = "blue", high = "red", mid = "white", 
						   midpoint = 0, limit = c(0,1), space = "Lab", 
						   name="Significance \n (p_val)") +
	  geom_text(aes(label = round(value, 3))) +
	  theme_minimal()+ 
	  theme(axis.text.x = element_text(angle = 45, vjust = 1, 
									   size = 12, hjust = 1)) +
  coord_fixed()
	
	
	
	
	
	# ==================================================================================================
	# ==================================================================================================
	# D. Three-level hierarchical model with all studies included for adverse events
	# ==================================================================================================
	# ==================================================================================================

	# GENERAL
	# Load data
	dat_adverse_events 	<- read_excel(file.path(wdir, "results_adverse-events.xlsx"))
	dat_adverse_events 	<- dat_adverse_events %>% na_if("NA") %>% drop_na(ae1i, n1i) # drop NA values in ae1i
	dat_ae 				<- escalc(measure="IRLN", xi=as.numeric(ae1i), ti=as.numeric(n1i), add=1/2, to="only0", data=dat_adverse_events)


	# V_mat <- impute_covariance_matrix(vi = dat$vi, cluster=dat$study, r = .7)
	dat$study[52] <- 66 # MUST be fixed with a reasonable value

	res <- rma.mv(yi, vi, slab=dat$slab, # V=V_mat, TODO add V_mat and remove vi
					random = ~ factor(study) | factor(category), #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst*study_type, tdist=TRUE,
					data=dat_ae, verbose=TRUE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
