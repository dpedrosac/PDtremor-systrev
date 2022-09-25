# This is code to run meta analyses on the results from the systematic literature review
# For further details on the data extraction cf. worksheets in __included.xlsx and 
# code in analysis_dataframe.r 
# Code developed by David Pedrosa

# Version 1.6 # 2022-09-06, final changes to analyses with some minor modifications in the groups

# ==================================================================================================
## In case of multiple people working on one project, this helps to create an automatic script
username <- Sys.info()["login"]
if (username == "dpedr") {
	wdir <- "D:/Jottacloud/PDtremor_syst-review/"
} else if (username == "david") {
	wdir <- "/media/storage/Jottacloud/PDtremor_syst-review/"
}
setwd(wdir)

# ==================================================================================================
## Specify packages of interest and load them automatically if needed
packages = c("readxl", "dplyr", "plyr", "tibble", "tidyr", "stringr", "openxlsx", "rstatix", "ggpubr",
"metafor", "tidyverse", "clubSandwich", "multcomp", "reshape2", "latex2exp", "irr", "tableone") # packages needed

## Load or install necessary packages
package.check <- lapply(
	packages,
	FUN = function(x) {
		if (!require(x, character.only = TRUE))
			{
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
# A. Three-level hierarchical model including all studies for effect sizes and category as 2nd level
# ==================================================================================================

# Prepare data adding groups, an "order" and a category
# =================
dat_results <- read.csv(file.path(wdir, "results", "results_meta-analysis.csv"), # load extracted data 
stringsAsFactors = F,  encoding="UTF-8")

groups <- data.frame(	levodopa=c("levodopa", rep(NA, 8)), # adds groups to dataframe
						betablocker=c("propanolol", "nipradilol", rep(NA, 7)), # zuranolone is no beta-blocker, as in previous versions
						primidone=c("primidone", rep(NA, 8)),
						anticholinergics=c("trihexphenidyl", "benztropin", "clozapine", rep(NA,6 )),
						amantadine=c("amantadine", rep(NA, 8)),
						memantine=c("memantine", rep(NA, 8)), 
						dopamine_agonists=c("ropinirole", "rotigotine", "apomorphine", "pramipexole", 
						"piribedil", "bromocriptine", "pergolide", "cabergoline", 
						"dihydroergocriptin"), 
						mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)),
						gabaergic_medication=c("clonazepam", "zuranolone", "gabapentin", rep(NA,6 )),
						cannabis=c("cannabis", rep(NA, 8)), 
						budipine=c("budipine", rep(NA, 8)),
						zonisamide=c("zonisamide", rep(NA, 8)),
						adenosine_antagonist=c("theophylline", "caffeine", "adenosineA2a", rep(NA, 6)),
						botox=c("botulinum", rep(NA, 8))
					)

dat <- dat_results %>% mutate(category=NA, class=NA, order_no=NA, treatment_diff=NA) # add columns to work with
iter<- 0
for (k in colnames(groups)){ # assigns category to the different treatments
	iter <- iter + 1
	temp <- dat_results %>% rowid_to_column %>% 
	filter(if_any(everything(),~str_detect(., c(paste(groups[[k]][!is.na(groups[[k]])],collapse='|'))))) 
	dat$category[temp$rowid]<- colnames(groups)[iter]
	dat$order_no[temp$rowid]<- iter
}

sorted_indices <- dat %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
						count() %>% arrange(desc(order_no))# 

dat <- dat %>% drop_na(-c(class, treatment_diff)) # remove NA, except to those in class and treatment_diff columns
dat$study <- 1:dim(dat)[1] # add study ID in the data frame for estimating results later
dat <- dat %>% arrange(desc(order_no), treatment)
attr(dat$yi, 'slab')<- dat$slab
temp <- dat %>% rowid_to_column %>% filter(if_any(everything(),
				~str_detect(., c(paste(c(	"levodopa", "amantadine", "clozapine", "budipine", "cannabis", 
											"zonisamide", "primidone", "botox"), collapse='|'))))) 
dat$treatment_diff[setdiff(1:dim(dat)[1], temp$rowid)] = dat$treatment[setdiff(1:dim(dat)[1], temp$rowid)]

coal = TRUE
if (coal==TRUE) { # merges columns from "treatment" and "treatment_diff" where NA appear 
	dat <- dat %>% mutate(treatment_diff = coalesce(treatment_diff, treatment))
}

# Prepare contingency table for quality of life estimates (Table 2 in the manuscript)
# =================
dat <- dat %>% mutate(., qualsyst_cat = case_when(
	qualsyst <= .50 ~ "low",
	(qualsyst > .50 & qualsyst < .70) ~ "medium",
	qualsyst >= .70 ~ "high"))

dat %>%
	group_by(study_type, qualsyst_cat) %>%
	dplyr::summarise(count = n()) %>%
	spread(qualsyst_cat, count)

# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
# =================
xrows_temp 	<- 1:150
spacing 	<- 5 
start 		<- 3
xrows 		<- c()
headings 	<- c()

for (i in 1:dim(sorted_indices)[1]){
	vector_group 	<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
	xrows 			<- c(xrows, vector_group)
	headings 		<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
	start 			<- start+sorted_indices[i,2] + 4
} 

# Impute Variance Covariance matrix according to correlated effects (see above)
# =================
covariance_adaptation = TRUE # in case of studies with the same participants, covariance should be adopted cf. 
if (covariance_adaptation==TRUE) { # https://www.jepusto.com/imputing-covariance-matrices-for-multi-variate-meta-analysis/
	dat$study.id <- dat$study
	dat$study.id[which(dat$slab=="Koller et al. 1987.1")] <- dat$study[which(dat$slab=="Koller et al. 1987")]
	dat$study.id[which(dat$slab=="Brannan & Yahr (1995)")] <- dat$study[which(dat$slab=="Brannan & Yahr (1995)")][1]
	dat$study.id[which(dat$slab=="Sahoo et al. 2020")] <- dat$study[which(dat$slab=="Sahoo et al. 2020")][1]
	dat$study.id[which(dat$slab=="Navan et al. 2003.1")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][1]
	dat$study.id[which(dat$slab=="Navan et al. 2003")] <- dat$study[which(dat$slab=="Navan et al. 2003.2")][2]
	dat$study.id[which(dat$slab=="Friedman et al. 1997")] <- dat$study[which(dat$slab=="Friedman et al. 1997")][1]
	dat$study.id[which(dat$slab=="Zach et al. 2020")] <- dat$study[which(dat$slab=="Dirkx et al. 2019")][1]
}

V_mat 	<- impute_covariance_matrix(vi = dat$vi, cluster=dat$study, r = .7)

# Run random effects meta-analysis using some "moderators" (qualsyst scores and study_type)
# =================
res 	<- rma.mv(yi, slab=dat$slab, V=V_mat,
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
comp_allstudies2 <- anova(res, res_wo2)

# Three-level meta analysis with levels 1. studies and 2. categories (of agents);
# Results show sigma1: variance between effect sizes within studies and sigma2: values between subgroups (categories 
# of treatment). Moreover is shows the amount (if at all) to which effects are influenced by other factors (moderators). 
# Finally, to see whether there is significant heterogeneity in the groups, comparisons of models with and without one
# level are appended (sigma=c(NA,0) or sigma=c(0,NA)). 

# Run model diagnostics
# =================
# Export diagnostics (profile) on model with all studies/categories (Supplementary Figure 1)
svg(filename=file.path(wdir, "results", "supplFigure.profile_all_groups.v1.0.svg"))
par(mfrow=c(2,1))
profile(res, sigma2=1, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{1}$ (factor = category))'))
title("Profile Likelihood Plot for model including all subgroups", line = -1, outer = TRUE)
profile(res, sigma2=2, steps=50, main=TeX(r'(Profile plot for $\sigma^{2}_{2}$ (factor = category/study_type))'))
dev.off()

# Calculation of I^2 statistic
Wtot 	<- diag(1/res$vi)
Xtot 	<- model.matrix(res)
Ptot 	<- Wtot - Wtot %*% Xtot %*% solve(t(Xtot) %*% Wtot %*% Xtot) %*% t(Xtot) %*% Wtot
100 * sum(res$sigma2) + (sum(res$sigma2) + (res$k-res$p)/sum(diag(Ptot)))
100 * res$sigma2 / (sum(res$sigma2) + (res$k-res$p)/sum(diag(Ptot)))

# Start plotting data using forest plot routines from {metafor}-package
# =================
cex_plot <- 1 # plot size
y_lim <- start # use last "start" as y-limit for forest plot
y_offset <- .5# offset used to plot the header of the column
# svg(filename=file.path(wdir, "results", "meta_analysis.all_groups.v2.1.svg"))
if (Sys.info()['sysname']!="Linux") {windowsFonts("Arial" = windowsFont("TT Arial"))}

forest(res, xlim=c(-7, 4.6), at=c(-3, -2, -1, 0, 1, 2), #atransf=exp,
ilab=cbind(dat$ni, dat$treatment_diff, dat$study_type, 
sprintf("%.02f",  dat$qualsyst), 
paste0(formatC(weights(res), format="f", digits=1, width=4), "%")),
xlab="Standardised mean change [SMCR]",
ilab.xpos=c(-5, -4, 2.5, 3, 3.5), 
cex=cex_plot, ylim=c(-1, y_lim),
rows=xrows,
efac =c(.8), #offset=y_lim-y_offset, 
mlab=mlabfun("RE Model for All Studies", res),
font=4, header="Author(s) and Year")

### set font expansion factor (as in forest() above) and use a bold font
op <- par(font=2, mar = c(2, 2, 2, 2))

text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
pos=c(2, 4), offset=y_lim-y_offset)

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
text(-7, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.2, 
 c(sprintf("Levodopa (n = %s)", sum(dat$ni[dat$category=="levodopa"])),
sprintf("beta-blocker (n = %s)", sum(dat$ni[dat$category=="betablocker"])),
sprintf("Primidone (n = %s)", sum(dat$ni[dat$category=="primidone"])),
sprintf("Anticholinergics (n = %s)", sum(dat$ni[dat$category=="anticholinergics"])),
sprintf("Amantadine (n = %s)", sum(dat$ni[dat$category=="amantadine"])),
sprintf("Memantine (n = %s)", sum(dat$ni[dat$category=="memantine"])),
sprintf("Dopamine agonists (n = %s)", sum(dat$ni[dat$category=="dopamine_agonists"])), 
sprintf("MAO-inhibitor (n = %s)", sum(dat$ni[dat$category=="mao_inhibitors"])), 
sprintf("Gabaergic medication (n = %s)", sum(dat$ni[dat$category=="gabaergic_medication"])), 
sprintf("Cannabis (n = %s)", sum(dat$ni[dat$category=="cannabis"])),
sprintf("Budipine (n = %s)", sum(dat$ni[dat$category=="budipine"])),
sprintf("Zonisamide (n = %s)", sum(dat$ni[dat$category=="zonisamide"])),
sprintf("Adenosine antagonists (n = %s)", sum(dat$ni[dat$category=="adenosine_antagonist"])),
sprintf("Botulinum toxin (n = %s)", sum(dat$ni[dat$category=="botox"])))
)

### set par back to the original settings
par(op)
### fit random-effects model in all subgroups
res.lev <- rma(yi, vi, subset=(category=="levodopa"), data=dat)
res.bet <- rma(yi, vi, subset=(category=="betablocker"), data=dat)
res.pri <- rma(yi, vi, subset=(category=="primidone"), data=dat)
res.cho <- rma(yi, vi, subset=(category=="anticholinergics"), data=dat)
res.mem <- rma(yi, vi, subset=(category=="memantine"),  data=dat)
res.ama <- rma(yi, vi, subset=(category=="amantadine"),     data=dat)
res.dag <- rma(yi, vi, subset=(category=="dopamine_agonists"),  data=dat)
res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"),     data=dat)
res.gab <- rma(yi, vi, subset=(category=="gabaergic_medication"),  data=dat)
res.cbd <- rma(yi, vi, subset=(category=="cannabis"),  data=dat)
res.bud <- rma(yi, vi, subset=(category=="budipine"),  data=dat)
res.zon <- rma(yi, vi, subset=(category=="zonisamide"),  data=dat)
res.ade <- rma(yi, vi, subset=(category=="adenosine_antagonist"),  data=dat)
res.bot <- rma(yi, vi, subset=(category=="botox"),  data=dat)

par(font=3)
yshift = .2
fac_cex = 1
if (Sys.info()['sysname']!="Linux") {windowsFonts("Arial" = windowsFont("TT Arial"))}

### add summary polygons for the three subgroups
addpoly(res.lev, row= 125.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.bet, row= 119.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.pri, row= 114.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.pri), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.cho, row= 103.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.ama, row= 97.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.ama), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.mem, row= 92.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.mem), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.dag, row= 58.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.dag), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.mao, row= 42.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.gab, row= 35.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.gab), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.cbd, row= 30.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.cbd), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.bud, row= 22.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.bud), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.zon, row= 16.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.zon), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.ade, row= 9.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.ade), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.bot, row= 1.5+yshift, mlab=mlabfun("RE Model for Subgroup", res.bot), cex=cex_plot*fac_cex, efac = c(.5))

abline(h=0)

### fit meta-regression model to test for subgroup differences
res_sg <- rma(yi, vi, mods = ~ category, data=dat)

### add text for the test of subgroup differences
text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
"Q[M]", " = ", .(formatC(res_sg$QM, digits=2, format="f")), ", df = ", .(res_sg$p - 1),
", p = ", .(formatC(res_sg$QMp, digits=3, format="f")))))
addpoly(res_sg, row= -1.8, mlab=mlabfun("RE Model for Subgroup", res_sg), cex=cex_plot*fac_cex, efac = c(.5))
if (Sys.info()['sysname']!="Linux") {windowsFonts("Arial" = windowsFont("TT Arial"))}
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

dat$residuals<- residuals.rma(res)
dat$precision<- 1/sqrt(dat$vi)
egger_res.all<- rma.mv(residuals~precision,vi,data=dat, random = ~ factor(study) | factor(category))

# ==================================================================================================
# B. Two-level hierarchical model with studies using dopamine agonists (only RCTs)
# ==================================================================================================

# Prepare data adding groups, an "order" and a category
# =================
dat_agonistsRCT <- dat %>% filter(category %in% c("levodopa", "dopamine_agonists")) %>% filter(study_type=="RCT")
remove_ergot_derivates=TRUE
if (remove_ergot_derivates){ # removes ergot derivatives as they are not marketed anymore
	dat_agonistsRCT <- dat_agonistsRCT %>% filter(treatment_diff!="dihydroergocriptin") %>% 
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

sorted_indices <- dat_agonistsRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
	count() %>% arrange(desc(order_no))# 


dat 		<- dat %>% drop_na(-c(class, treatment_diff)) 	# remove NA, except those in "class"-/"treatment_diff"-columns
dat$study 	<- 1:dim(dat)[1] 								# add study ID to data frame for estimating results later
dat	 		<- dat %>% arrange(desc(order_no), treatment)
attr(dat$yi, 'slab')<- dat$slab
temp 		<- dat %>% rowid_to_column %>% filter(if_any(everything(),
				~str_detect(., c(paste(c(	"levodopa", "amantadine", "clozapine", "budipine", "cannabis", 
											"zonisamide", "primidone", "botox"), collapse='|'))))) 
dat$treatment_diff[setdiff(1:dim(dat)[1], temp$rowid)] = dat$treatment[setdiff(1:dim(dat)[1], temp$rowid)]

coal = TRUE
if (coal==TRUE) { # merges columns from "treatment" and "treatment_diff" where NA appear 
	dat <- dat %>% mutate(treatment_diff = coalesce(treatment_diff, treatment))
}

# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
# =================
xrows_temp 	<- 1:150
spacing 	<- 5 
start 		<- 3
xrows 		<- c()
headings 	<- c()

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
res_agonistsRCT		<- rma.mv(yi, slab=dat_agonistsRCT$slab, V=V_mat_agonistsRCT,
							random = ~1 | as.factor(treatment)/study, #~ 1 | slab, #/factor(study_type)# , 
							mods= ~ qualsyst,tdist=TRUE, #struct="DIAG",
							data=dat_agonistsRCT, verbose=TRUE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
res_agonistsRCT_wo1 <- rma.mv(yi, slab=dat_agonistsRCT$slab, V=V_mat_agonistsRCT,
					random = ~1 | as.factor(treatment)/study, 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG", #qualsyst*study_type
					data=dat_agonistsRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(0, NA))
res_agonistsRCT_wo2 <- rma.mv(yi, slab=dat_agonistsRCT$slab, V=V_mat_agonistsRCT,
					random = ~1 | as.factor(treatment)/study, 
					mods= ~ qualsyst, tdist=TRUE, struct="DIAG", #qualsyst*study_type
					data=dat_agonistsRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(NA, 0))
comp_agonistsRCT1 <- anova(res_agonistsRCT, res_agonistsRCT_wo1)
comp_agonistsRCT2 <- anova(res_agonistsRCT, res_agonistsRCT_wo2)


# Three-level meta analysis with levels studies and subgroups (of agonists);
# Results show sigma1: variance between effect sizes within studies and sigma2: values between subgroups (categories 
# of dopamine agonists). Moreover is shows the amount (if at all) to which effects are influenced by other factors (moderators). 
# Finally, to see whether there is significant heterogeneity in the groups, comparisons of models with and without one
# level are appended (sigma=c(NA,0) or sigma=c(0,NA)). 

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
100 * sum(res_agonistsRCT$sigma2) + (sum(res_agonistsRCT$sigma2) + (res_agonistsRCT$k-res_agonistsRCT$p)/sum(diag(PagoRCT)))
100 * res_agonistsRCT$sigma2 / (sum(res_agonistsRCT$sigma2) + (res_agonistsRCT$k-res_agonistsRCT$p)/sum(diag(PagoRCT)))

# Start plotting data using forest plot routines from {metafor}-package
# =================
cex_plot 	<- 1 # plot size
y_lim 		<- start # use last "start" as y-limit for forest plot
y_offset 	<- .5# offset used to plot the header of the column
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
	font=4, header="Author(s) and Year"
	)

### set font expansion factor (as in forest() above) and use a bold font
op <- par(font=2, mar = c(2, 2, 2, 2))

text((c(-.5, .5)), cex=cex_plot, y_lim, c("Tremor reduction", "Tremor increase"), 
pos	<- c(2, 4), offset=y_offset)

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
							test = adjusted("BH")) # pairwise comparisons
mat2plot				<- matrix(0, nrow = 5, ncol = 5, dimnames=list(toupper(substr(groups, 1, 3)), toupper(substr(groups, 1, 3)))) + diag(5)/1000000
mat_significance 		<- mat2plot 
mat2plot[lower.tri(mat2plot, diag = FALSE)] 		<- unname(pwc$test$tstat)
mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
melted_cormat 			<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
melted_cormat$pval 		<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
melted_cormat$pval 		<- melted_cormat$pval$value # dodgy solution
colnames(melted_cormat)	<- c("Var1", "Var2", "tstat", "pval")
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
# C. Two-level hierarchical model with studies using mao-inhibitors (only RCTs)
# ==================================================================================================

# Prepare data adding groups, an "order" and a category
# =================
groups 		<- data.frame(mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)))
dat_maoRCT 	<- dat_results %>% mutate(category=NA, class=NA, order_no=NA, treatment_diff=NA) # add columns to dataframe

iter<- 0
for (k in colnames(groups)){ # assigns category to the different treatments
	iter <- iter + 1
	temp <- dat_results %>% rowid_to_column %>% 
	filter(if_any(everything(),~str_detect(., c(paste(groups[[k]][!is.na(groups[[k]])],collapse='|'))))) 
	dat_maoRCT$category[temp$rowid]<- colnames(groups)[iter]
	dat_maoRCT$order_no[temp$rowid]<- iter
}

sorted_indices <- dat_maoRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
					count() %>% arrange(desc(order_no))# 

dat_maoRCT 			<- dat_maoRCT %>% drop_na(-c(class, treatment_diff)) # get data of interest, that
dat_maoRCT$study	<- 1:dim(dat_maoRCT)[1] # add the study ID in the daat frame
dat_maoRCT			<- dat_maoRCT %>% arrange(desc(order_no), treatment)
attr(dat_maoRCT$yi, 'slab')<- dat_maoRCT$slab
temp 				<- dat_maoRCT %>% rowid_to_column %>% 
						filter(if_any(everything(),
							~str_detect(., c(paste(c("rasagiline", "selegiline", "safinamide")), collapse='|'))))) 
dat_maoRCT$treatment_diff[setdiff(1:dim(dat_maoRCT)[1], temp$rowid)] <- dat_maoRCT$treatment[setdiff(1:dim(dat_maoRCT)[1], temp$rowid)]

coal = TRUE
if (coal==TRUE) { # merges columns from "treatment" and "treatment_diff" where NA appear 
	dat_maoRCT <- dat_maoRCT %>% mutate(treatment_diff = coalesce(treatment	_diff, treatment))
}

groups 	<- c("selegiline", "safinamide", "rasagiline")
iter	<- 0
for (i in groups){ 
	iter <- iter + 1
	temp <- dat_maoRCT %>% rowid_to_column %>% 
	filter(if_any(everything(),~str_detect(., i))) 
	dat_maoRCT$order_no[temp$rowid]<- iter
}

sorted_indices <- dat_maoRCT %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
					count() %>% arrange(desc(order_no))# 

# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
# =================
xrows_temp 	<- 1:150
spacing 	<- 5 
start 		<- 3
xrows 		<- c()
headings 	<- c()

for (i in 1:dim(sorted_indices)[1]){
	vector_group <- xrows_temp[start:(start+sorted_indices[i,2]-1)]
	xrows <- c(xrows, vector_group)
	headings <- c(headings, start+sorted_indices[i,2]-1 + 1.5)
	start <- start+sorted_indices[i,2] + 4
} 

# Impute Variance Covariance matrix according to correlated effects (see above)
# =================
V_mat_maoRCT <- impute_covariance_matrix(vi = dat_maoRCT$vi, cluster=dat_maoRCT$study, r = .7)


res_maoRCT		<- rma.mv(yi, slab=dat_maoRCT$slab, V=V_mat_maoRCT,
					random = ~1 | as.factor(treatment)/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_maoRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"))
res_maoRCT_wo1	<- rma.mv(yi, slab=dat_maoRCT$slab, V=V_mat_maoRCT,
					random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_maoRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(0, NA))
res_maoRCT_wo2	<- rma.mv(yi, slab=dat_maoRCT$slab, V=V_mat_maoRCT,
					random = ~1 | treatment/study, #~ 1 | slab, #/factor(study_type)# , 
					mods= ~ qualsyst, tdist=TRUE, #struct="DIAG",
					data=dat_maoRCT, verbose=FALSE, control=list(optimizer="optim", optmethod="Nelder-Mead"), sigma=c(NA, 0))
comp_maoRCT1 <- anova(res_maoRCT, res_maoRCT_wo1)
comp_maoRCT2 <- anova(res_maoRCT, res_maoRCT_wo2)


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
cex_plot <- 1 # plot size
y_lim <- start # use last "start" as y-limit for forest plot
y_offset <- .5# offset used to plot the header of the column
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
groups <- c("selegiline", "safinamide", "rasagiline")
levels(dat_maoRCT) <- groups
res_overall_maoRCT <- rma(yi, vi, mods = ~ treatment_diff - 1, data=dat_maoRCT)

### add text for the test of subgroup differences
text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
 "Q[M]", " = ", .(formatC(res_overall_maoRCT$QM, digits=2, format="f")), ", df = ", .(res_overall_maoRCT$p - 1),
 ", p = ", .(formatC(res_overall_maoRCT$QMp, digits=3, format="f")))))
dev.off()

# Run post-hoc analyses
# =================
pwc <- summary(glht(res_overall_maoRCT, linfct=cbind(contrMat(rep(1,3), type="Tukey"))), 
test=adjusted("BH")) # pairwise comparisons
mat2plot <- matrix(0, nrow = 3, ncol = 3, dimnames=list(toupper(substr(groups, 1, 3)), toupper(substr(groups, 1, 3)))) + + diag(3)/1000000
mat_significance <- mat2plot 
mat2plot[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$tstat)
mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
melted_cormat <- melt(mat2plot, na.rm = TRUE) # converts to "long" format
melted_cormat$pval <- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
melted_cormat$pval <- melted_cormat$pval$value # dodgy solution
colnames(melted_cormat)<-c("Var1", "Var2", "tstat", "pval")
melted_cormat[melted_cormat==0] <- NA # remove zero values and convert to NA
melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001] <- "*"
melted_cormat$pval[melted_cormat$pval<.001] <- "**"
melted_cormat$pval[melted_cormat$pval>.05] <- NA

# Create a heatmap for pairwise comparisons
svg(filename=file.path(wdir, "results", "heatmaps.maoRCT.v1.0.svg"))
windowsFonts("Arial" = windowsFont("TT Arial"))
ggplot(data = melted_cormat %>% drop_na(-pval), aes(Var2, Var1, fill = tstat)) +
geom_tile(color = "white") +
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
# D. Three-level hierarchical model with all studies included for adverse events
# ==================================================================================================

# Load and prepare data
# =================
dat_aes 				<- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
do_stats				<- mean(100*as.numeric(dat_aes$do2i)/as.numeric(dat_aes$n2i), na.rm=T) # adapt according to needed value 		
dat_aes 				<- dat_aes %>% dplyr::filter(treatment2!="pergolide") %>% na_if("NA") %>% drop_na(ae1i, ae2i, n1i, n2i) %>% 
						mutate_at(c(	'ae1i', 'ae2i','do1i', 'do2i', 'n1i', 'n2i', 
										'age1', 'age2', 'updrs1', 'updrs2'), as.numeric)# drop NA values in ae1i convert all necessary columns to numeric
dat_aes					<- dat_aes %>% dplyr::filter(.,study_type=="RCT")
dat_aes$study.id		<- 1:dim(dat_aes)[1]
attr(dat_aes$yi, 'slab')<- dat_aes$slab

dat_aes$order_no[dat_aes$order_no==5] <- 1 # order is changed manually
dat_aes$order_no[dat_aes$order_no==6] <- 2
dat_aes$order_no[dat_aes$order_no==7] <- 3
dat_aes$order_no[dat_aes$order_no==8] <- 4
dat_aes$order_no[dat_aes$order_no==10]<- 5
dat_aes$order_no[dat_aes$order_no==12]<- 6

# dat_aes 				<- dat_aes %>% drop_na(do1i, do2i) not needed at this point
dat_aes$study.id 		<- 1:dim(dat_aes)[1]

sorted_indices 			<- dat_aes %>% dplyr::select(order_no) %>% drop_na() %>% gather(order_no) %>% 
							count() %>% arrange(desc(order_no))# 

# Prepare additional information necessary to plot data correctly, i.e. distances between subgroups etc.
# =================
xrows_temp 	<- 1:150
spacing 	<- 5 
start 		<- 3
xrows 		<- c()
headings 	<- c()

for (i in 1:dim(sorted_indices)[1]){
	vector_group<- xrows_temp[start:(start+sorted_indices[i,2]-1)]
	xrows 		<- c(xrows, vector_group)
	headings 	<- c(headings, start+sorted_indices[i,2]-1 + 1.5)
	start 		<- start+sorted_indices[i,2] + 4
} 

dat_aes <- escalc(measure="RR", ai=dat_aes$ae1i, 
					bi=dat_aes$n1i - dat_aes$ae1i, 
					ci=dat_aes$ae2i, 
					di=dat_aes$n2i - dat_aes$ae2i,
					n1i=n1i, n2i=n2i,
					add=1/2, to="all", data=dat_aes)

# Impute variance-covariance-matrix according to correlated effects (see above)
# =================
V_matAE <- impute_covariance_matrix(vi = dat_aes$vi, cluster=dat_aes$study, r = .7)

fit_ae 	<- rma.mv(yi, slab=dat_aes$slab, V=V_matAE, random = ~ 1 | as.factor(category)/study.id,
					data=dat_aes, mods= ~ qualsyst, method="REML")

cex_plot<- 1 # plot size
y_lim 	<- 50 # use last "start" as y-limit for forest plot
y_offset<- .5# offset used to plot the header of the column
# svg(filename=file.path(wdir, "results", "meta_analysis.agonistsRCT.v1.0.svg"))
windowsFonts("Arial" = windowsFont("TT Arial"))
forest(fit_ae, xlim=c(-9, 10), at=log(c(0.25, 1, 4, 20)), atransf=exp,
		ilab=cbind(dat_aes$ni, dat_aes$treatment, dat_aes$ae2i, dat_aes$ae1i, dat_aes$study_type, 
		sprintf("%.02f",  dat_aes$qualsyst), 
		paste0(formatC(weights(fit_ae), format="f", digits=1, width=4), "%")),
		xlab="Relative risk",
		ilab.xpos=c(-5.5, -4, -3, -2, 5, 6, 7), 
		cex=cex_plot, ylim=c(-1, y_lim),
		rows=xrows,
		efac =c(.8), #offset=y_offset, 
		mlab=mlabfun("RE Model for All Studies", res_agonistsRCT),
		font=4, header="Author(s) and Year")

### set font expansion factor (as in forest() above) and use a bold font
op <- par(font=2, mar = c(2, 2, 2, 2))

pos=c(2, 4), offset=y_offset)

### Add column headings for everything defined in "forest::ilab"
text(c(-5.5), cex=cex_plot, y_lim, c("n"))
text(c(-4), cex=cex_plot, y_lim, c("agent"))
text(c(-3), cex=cex_plot, y_lim, c("AEctrl"))
text(c(-2), cex=cex_plot, y_lim, c("AEver"))
text(c(5), cex=cex_plot, y_lim, c("type"))
text(c(6), cex=cex_plot, y_lim+2*y_offset, c("QualSyst"))
text(c(6), cex=cex_plot, y_lim, c("score"))
text(c(7), cex=cex_plot, y_lim, c("weight"))

### add text for subgroups
### add text for the subgroups
text(-9, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot*1.15, c(
   sprintf("Dopamine agonists (n = %s)", sum(dat_aes$ni[dat_aes$category=="dopamine_agonists"])),
   sprintf("Clozapine (n = %s)", sum(dat_aes$ni[dat_aes$category=="clozapine"])), 
   sprintf("MAO-inhibitor (n = %s)", sum(dat_aes$ni[dat_aes$category=="mao_inhibitors"])), 
   sprintf("Budipine (n = %s)", sum(dat_aes$ni[dat_aes$category=="budipine"])),
   sprintf("Adenosine antagonists (n = %s)", sum(dat_aes$ni[dat_aes$category=="adenosine_antagonist"]))
))

### set par back to the original settings
par(op)

### fit random-effects model in all subgroups
res.dop <- rma(yi, vi, subset=(category=="dopamine_agonists"), data=dat_aes)
res.clo <- rma(yi, vi, subset=(category=="clozapine"), data=dat_aes)
res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"), data=dat_aes)
res.bud <- rma(yi, vi, subset=(category=="budipine"), data=dat_aes)
res.ade <- rma(yi, vi, subset=(category=="adenosine_antagonist"), data=dat_aes)

yshift = .2
fac_cex = 1
### add summary polygons for the three subgroups
addpoly(res.ade, atransf=exp, row= 1.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.ade), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.bud, atransf=exp, row= 6.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.bud), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.mao, atransf=exp, row= 11.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.clo, atransf=exp, row= 25.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=cex_plot*fac_cex, efac = c(.5))
addpoly(res.dop, atransf=exp, row= 30.5 + yshift, mlab=mlabfun("RE Model for Subgroup", res.dop), cex=cex_plot*fac_cex, efac = c(.5))

abline(h=0)

### fit meta-regression model to test for subgroup differences
dat_aes$category <- as.factor(dat_aes$category)
groups <- c("adenosine_antagonist", "budipine", "mao_inhibitors", "clozapine", "dopamine_agonists")
levels(dat_aes) <- groups
res_overallAE <- rma(yi, vi, mods = ~ category - 1, data=dat_aes)

### add text for the test of subgroup differences
text(-9, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
 "Q[M]", " = ", .(formatC(res_overallAE$QM, digits=2, format="f")), ", df = ", .(res_overallAE$p - 1),
 ", p = ", .(formatC(res_overallAE$QMp, digits=3, format="f")))))
dev.off()

# Run post-hoc analyses
# =================
pwc 				<- summary(glht(res_overallAE, linfct=cbind(contrMat(rep(1,5), type="Tukey"))), 
						test=adjusted("BH")) # pairwise comparisons
mat2plot 			<- matrix(0, nrow = 5, ncol = 5, dimnames=list(toupper(substr(groups, 1, 5)), toupper(substr(groups, 1, 5)))) + diag(5)/1000000
mat_significance 	<- mat2plot 
mat2plot[lower.tri(mat2plot, diag = FALSE)] 		<- unname(pwc$test$tstat)
mat_significance[lower.tri(mat2plot, diag = FALSE)] <- unname(pwc$test$pvalues)
melted_cormat 		<- melt(mat2plot, na.rm = TRUE) # converts to "long" format
melted_cormat$pval 	<- melt(mat_significance, na.rm = TRUE) %>% dplyr::select(value)
melted_cormat$pval 	<- melted_cormat$pval$value # dodgy solution ; )
colnames(melted_cormat)<-c("Var1", "Var2", "tstat", "pval")

melted_cormat[melted_cormat==0] 									<- NA # remove zero values and convert to NA
melted_cormat$pval[melted_cormat$pval<.05 & melted_cormat$pval>.001]<- "*"
melted_cormat$pval[melted_cormat$pval<.001] 						<- "**"
melted_cormat$pval[melted_cormat$pval>.05] 							<- NA

# Create a heatmap for pairwise comparisons
# svg(filename=file.path(wdir, "results", "heatmaps.aesRCT.v1.0.svg"))
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

dat_table_one <- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
dat_table_one <- dat_table_one %>% mutate_at(c('age1', 'age2', 'n1i', 'n2i', 'updrs1', 'updrs2', 'qualsyst'), as.numeric) %>% 
mutate(age_collated = (n1i*age1 + n2i*age2)/(n1i + n2i))
dat_table_one$age_collated <- ifelse(is.na(dat_table_one$age_collated), dat_table_one$age1, dat_table_one$age_collated)

dat_table_one <- dat_table_one %>% mutate(updrs_collated = (n1i*updrs1 + n2i*updrs2)/(n1i + n2i))
dat_table_one$updrs_collated<- ifelse(is.na(dat_table_one$updrs_collated), dat_table_one$updrs1, dat_table_one$updrs_collated)
dat_table_one$study_type <- as.factor(dat_table_one$study_type)
dat_table_one$category <- as.factor(dat_table_one$category)

dat_table_one$category[which(dat_table_one$category=="clozapine")] = "anticholinergics"
NumVars <- c("age_collated", "qualsyst", "updrs_collated", "ni")
catVars <- c("study_type", "category")
myVars <- c(catVars, NumVars)
tableOne <- CreateTableOne(vars = myVars, data = dat_table_one, factorVars = catVars)
write.csv(print(tableOne), file=file.path(getwd(), "results", "tableOne.v1.0.csv"))


# ==================================================================================================
# F. Intraclass correlation between Qualsyst scores
# ==================================================================================================

results_file 	<-  file.path(wdir, "__included_studies.xlsx")
excel_sheets(path = results_file) 								# Names of sheets for later use
tab_names 		<- excel_sheets(path = results_file)


# ==================================================================================================
# Extract authors and years from worksheets
authors 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('Abbr', 'Author'), range = "C2:D2"))

years 		<- lapply(tab_names, function(x) read_excel(results_file, 															
															sheet = x, col_names=c('Abbr', 'Year'), range = "C4:D4"))
authors_wide<- data.frame(data.table::rbindlist(authors, idcol='ID')) %>% dplyr::select(Author)
years_short <- data.frame(data.table::rbindlist(years, idcol='ID')) %>% dplyr::select(Year)

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
# Extract Qualsyst scores from all worksheets
qualsyst 	<- lapply(tab_names, function(x) read_excel(results_file, 
															sheet = x, col_names=c('Abbr', 'QS1.1', 'QS1.2', 'QS2.1', 'QS2.2'), range = "B69:F69"))
qualsyst_short <- data.frame(data.table::rbindlist(qualsyst, idcol='ID'))
dat_icc <- tibble(qualsyst_short) %>% drop_na(!Abbr) %>% mutate(rater1=QS1.1/QS1.2, rater2=QS2.1/QS2.2)

icc(
  dat_icc %>% dplyr::select(rater1, rater2), model = "twoway", 
  type = "agreement", unit = "single"
  )


# ==================================================================================================
# ==================================================================================================
# G. Anova for subgroups of different agents
# ==================================================================================================
# ==================================================================================================

# Load data and extract further information needed (e.g., total age, cumulative updrs between groups, etc.)
# =================
dat_demographics <- read_excel(file.path(wdir, "results_adverse-events.v2.0.xlsx"))
dat_demographics <- dat_demographics %>% mutate_at(c('age1', 'age2', 'n1i', 'n2i', 'updrs1', 'updrs2', 'qualsyst'), as.numeric) %>% 
	mutate(age_collated = (n1i*age1 + n2i*age2)/(n1i + n2i))
dat_demographics$age_collated <- ifelse(is.na(dat_demographics$age_collated), dat_demographics$age1, dat_table_one$age_collated)
dat_demographics <- dat_demographics %>% mutate(updrs = (n1i*updrs1 + n2i*updrs2)/(n1i + n2i))
dat_demographics$updrs<- ifelse(is.na(dat_demographics$updrs), dat_demographics$updrs1, dat_demographics$updrs)

# Run ANOVA for age and UPDRS (Supplementary Figure 2 in the manuscript)
# =================
stat.test.age <- dat_demographics %>% anova_test(age_collated ~ category)
post_hoc.age <- aov(age_collated ~ category, data=dat_demographics) %>% tukey_hsd()

svg(filename=file.path(wdir, "results", "group_comparisons.v1.0.svg"))
age <- ggboxplot(dat_demographics, x = "category", y = "age_collated") +
stat_compare_means(method = "anova", label.y = 75)+ 
xlab("") + ylim(55, 75) +
  theme(axis.ticks.x = element_blank(),
        axis.text.x = element_blank(),
		panel.grid.major.y = element_blank()) + 
geom_hline(yintercept = mean(dat_demographics$age_collated, na.rm=TRUE), linetype = 2)

summary(aov(updrs~category, dat_demographics %>% drop_na(updrs)))
stats_updrs <- aov(updrs~category, dat_demographics %>% 
	drop_na(updrs)) %>% tukey_hsd() %>%
	mutate(y.position = c(rep(0, 32), 62, rep(0, 3)))
LabelsAgents <- c(	'Botulinum toxin', 
					'Adenosine antagonists',
					'Zonisamide', 
					'Budipine', 
					'Cannabis', 
					'Gabaergic agents', 
					'MAO-inhibitors', 
					'Dopamine agonists', 
					'Memantine', 
					'Amantadine', 
					'Anticholinergics', 
					'Primidone', 
					'Beta-blocker', 
					'Levodopa')
updrs <- ggboxplot(dat_demographics, x = "category", y = "updrs") +
xlab("") + ylim(0, 65) +
rotate_x_text(angle = 45) +
  stat_pvalue_manual(
    data = stats_updrs, label = "p.adj.signif",
    y.position = "y.position", hide.ns = TRUE
    ) +
scale_x_discrete(labels= LabelsAgents) + 
# stat_compare_means(method = "t.test", label.y = 60, aes(label=..p.adj..)) + 
geom_hline(yintercept = mean(dat_demographics$updrs, na.rm=TRUE), linetype = 2)

grid.arrange(age, updrs, nrow=2)
# plot one below the other
dev.off()
