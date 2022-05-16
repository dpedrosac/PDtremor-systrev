# This is code to run meta analyses on the results from the systematic literature review
# For furtehr details on the data extraction cf. worksheets in __included.xlsx and 
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
				"metafor", "tidyverse", "clubSandwich") 											# packages needed

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
# ==================================================================================================
# Three-level hierarchical model with all studies included
# ==================================================================================================
# ==================================================================================================

# GENERAL
# Load data
dat_results <- read.csv(file.path(wdir, "results", "results_meta-analysis.csv"), stringsAsFactors = F,  encoding="UTF-8")
cex_plot 	<- .8
groups 		<- data.frame(	levodopa=c("levodopa", rep(NA, 8)), # adds groups to the dataframe
							betablocker=c("propanolol", "zuranolone", "nipradilol", rep(NA, 6)),
						    primidone=c("primidone", rep(NA, 8)),
							anticholinergics=c("trihexphenidyl", "benztropin", rep(NA,7 )),
							amantadine=c("amantadine", rep(NA, 8)), 
							dopamine_agonists=c("ropinirole", "rotigotine", "apomorphine", "pramipexole", "piribedil", "bromocriptine", "pergolide", "cabergoline", "dihydroergocriptin"), 
							clozapine=c("clozapine", rep(NA, 8)), 
							mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 6)),
							cannabis=c("cannabis", rep(NA, 8)), 
							budipine=c("budipine", rep(NA, 8)),
							zonisamide=c("zonisamide", rep(NA, 8)),
							adenosine_antagonist=c("theophylline", "caffeine", "adenosineA2a", rep(NA, 6)),
						  	botox=c("botulinum", rep(NA, 8)))
dat 		<- dat_results %>% mutate(category=NA, class=NA, order=NA) # add columns to dataframe to work with

iter 	<- 0
for (k in colnames(groups)){ # assigns category to the different treatments
	iter 					<- iter +1
	comparator_groups 		<- groups[[k]][!is.na(groups[[k]])]
	temp 					<- dat_results %>% rowid_to_column %>% 
							filter(if_any(everything(),~str_detect(., c(paste(comparator_groups,collapse='|'))))) 
	dat$category[temp$rowid]<- colnames(groups)[iter]
	dat$order[temp$rowid] 	<- iter
}

sorted_indices 	<- dat %>% select(order) %>% drop_na() %>% gather(order) %>% count() %>% arrange(desc(order))# 
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

dat <- dat %>% drop_na(-class)
dat$study <- 1:dim(dat)[1]
#dat <- dat %>% arrange(-order, class, -slab)
dat <- dat %>% arrange(desc(order), treatment)
attr(dat$yi, 'slab') <- dat$slab
temp 				<- dat %>% rowid_to_column %>% 
							filter(if_any(everything(),
									~str_detect(., c(paste(c("levodopa", "amantadine", "clozapine", 
																"budipine", "cannabis", "zonisamide", "primidone", 																			"botox"),collapse='|'))))) 
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


# Imput the Variance Covariance matrix according to correlated effects (see above
V_mat <- impute_covariance_matrix(vi = dat$vi, cluster=dat$study, r = .7)

res <- rma.mv(yi, slab=dat$slab, V=V_mat,
				random = ~ factor(study) | factor(category), #~ 1 | slab, #/factor(study_type)# , 
				mods= ~ qualsyst*study_type, tdist=TRUE,
				data=dat, verbose=TRUE, control=list(optimizer="optim", optmethod="Nelder-Mead"))

# par(mfrow=c(2,1))
# profile(res, tau2=1)
# profile(res, rho=1)

# Start plotting data using a forest plot
forest(res, xlim=c(-10, 4.6), at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
       ilab=cbind(sprintf("%.02f",  dat$qualsyst), dat$ni, 
	   paste0(formatC(weights(res), format="f", digits=1, width=4), "%"),
	   dat$treatment2, dat$study_type), #, dat$tneg, dat$cpos, dat$cneg),
       ilab.xpos=c(2, -5, 2.5, -4, 1.5), 
	   #ilab.xpos=c(-9.5,-8,-6,-4.5), 
	   cex=cex_plot, ylim=c(-1, y_lim),
       rows=xrows, #c(3:4, 9:18, 23:27, 32:64, 69, 74:81),
	   #slab=dat$slab, 
	   #order=dat$order,
       mlab=mlabfun("RE Model for All Studies", res),
       font=4, header="Author(s) and Year")
 
### set font expansion factor (as in forest() above) and use a bold font
op <- par(cex=0.5, font=4, mar = c(6, 6, 6, 6))
 
### add additional column headings to the plot
text(c(2), y_lim, c("QualSyst"))
text(c(2), y_lim-1, c("score"))
text(c(2.5), y_lim-1, c("weight"))
text(c(-5), y_lim-1, c("n"))
text(c(-4), y_lim-1, c("agent"))
text(c(1.5), y_lim-1, c("type"))
 
### switch to bold italic font
par(font=4)
 
### add text for the subgroups
text(-10, sort(headings, decreasing=TRUE), pos=4, cex=cex_plot, c(sprintf("Levodopa (n = %s)", sum(dat$ni[dat$category=="levodopa"])),
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

### add summary polygons for the three subgroups
addpoly(res.lev, row= 114.5, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=cex_plot*.8, efac = c(.8))
addpoly(res.bet, row= 107.5, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*.8, efac = c(.8))
addpoly(res.pri, row= 102.5, mlab=mlabfun("RE Model for Subgroup", res.pri), cex=cex_plot*.8, efac = c(.8))
addpoly(res.cho, row= 96.5, mlab=mlabfun("RE Model for Subgroup", res.bet), cex=cex_plot*.8, efac = c(.8))
addpoly(res.ama, row= 90.5, mlab=mlabfun("RE Model for Subgroup", res.ama), cex=cex_plot*.8, efac = c(.8))
addpoly(res.clo, row= 57.5, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=cex_plot*.8, efac = c(.8))
addpoly(res.dag, row= 48.5, mlab=mlabfun("RE Model for Subgroup", res.dag), cex=cex_plot*.8, efac = c(.8))
addpoly(res.mao, row= 33.5, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=cex_plot*.8, efac = c(.8))
addpoly(res.cbd, row= 28.5, mlab=mlabfun("RE Model for Subgroup", res.cbd), cex=cex_plot*.8, efac = c(.8))
addpoly(res.bud, row= 21.5, mlab=mlabfun("RE Model for Subgroup", res.bud), cex=cex_plot*.8, efac = c(.8))
addpoly(res.zon, row= 15.5, mlab=mlabfun("RE Model for Subgroup", res.zon), cex=cex_plot*.8, efac = c(.8))
addpoly(res.ade, row= 9.5, mlab=mlabfun("RE Model for Subgroup", res.ade), cex=cex_plot*.8, efac = c(.8))
addpoly(res.bot, row= 1.5, mlab=mlabfun("RE Model for Subgroup", res.bot), cex=cex_plot*.8, efac = c(.8))
 
abline(h=0)

### fit meta-regression model to test for subgroup differences
res_sg <- rma(yi, vi, mods = ~ category, data=dat)
 
### add text for the test of subgroup differences
text(-10, -1.8, pos=4, cex=cex_plot, bquote(paste("Test for Subgroup Differences: ",
     "Q[M]", " = ", .(formatC(res_sg$QM, digits=2, format="f")), ", df = ", .(res_sg$p - 1),
     ", p = ", .(formatC(res_sg$QMp, digits=3, format="f")))))
# addpoly(res_sg, row= -1.8, mlab=mlabfun("RE Model for Subgroup", res_sg), cex=cex_plot)



