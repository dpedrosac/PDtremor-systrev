# This is code to run meta analyses on the results from the systematic literature review
# For furtehr details on the data extraction cf. worksheets in __included.xlsx and 
# code in analysis_dataframe.r 
# Code developed by David Pedrosa

# Version 1.0 # 2022-03-27, added separate file for meta analyses

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
# Load data
dat_results <- read_csv(file.path(wdir, "results", "results_meta-analysis.csv"))

# Subset clozapine as example
dat <- dat_results %>% filter(if_any(everything(),~str_detect(., c("clozapine|clozapine")))) 
res <- rma(yi, vi, data=dat, mods= ~ qualsyst, digits=2)
res$slab <- dat$slab

mlabfun <- function(text, res) {
   list(bquote(paste(.(text),
      " (Q = ", .(formatC(res$QE, digits=2, format="f")),
      ", df = ", .(res$k - res$p),
      ", p ", .(metafor:::.pval(res$QEp, digits=2, showeq=TRUE, sep=" ")), "; ",
      I^2, " = ", .(formatC(res$I2, digits=1, format="f")), "%, ",
      tau^2, " = ", .(formatC(res$tau2, digits=2, format="f")), ")")))}

weights <- paste0(formatC(weights(res), format="f", digits=1), "%")
### forest plot with extra annotations
forest(res, atransf=exp, at=log(c(.05, .25, 1, 4, 20)), xlim=c(-16,6),
       weights= dat_results$qualsyst,
	   ilab=cbind(dat$treatment, weights),
       ilab.xpos=c(-9.5,-6.5), cex=.75, header="Author(s) and Year",
       mlab=mlabfun("RE Model for All Studies", res), slab=res$slab)
op <- par(cex=.75, font=2)
text(c(-9.5,-6.5), dim(dat)[1]+2, c("Treatment", "Weight"))
#text(c(-8.75,-5.25),     16, c("Vaccinated", "Control"))
par(op)
abline(h=0)
sav <- predict(res, newmods = mean(dat$qualsyst))
addpoly(sav$pred, sei=sav$se, mlab="Adjusted Effect")


# ==================================================================================================
# ==================================================================================================
# Three-level hierarchical model with all studies included
# ==================================================================================================
# ==================================================================================================

# GENERAL
cex_plot 	<- .5
groups 		<- data.frame(	levodopa=c("levodopa", rep(NA, 6)), # adds groups to the dataframe
							amantadine=c("amantadine", rep(NA, 6)), 
							dopamine_agonists=c("ropinirole", "rotigotine", "apomorphine", "pramipexole", "piribedil", "bromocriptine", "pergolide"), 
							clozapine=c("clozapine", rep(NA, 6)), 
							mao_inhibitors=c("rasagiline", "selegiline", "safinamide", rep(NA, 4)),
							cannabis=c("cannabis", rep(NA, 6)), 
							botox=c("botulinum", rep(NA, 6)))
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
# ==================================================================================================


dat <- dat %>% drop_na(-class)
dat$study <- 1:dim(dat)[1]
dat <- dat %>% arrange(desc(order))
attr(dat$yi, 'slab') <- dat$slab

res <- rma(yi, vi, slab=dat$slab, data=dat) #	  

res <- rma.mv(yi, vi, slab=dat$slab, random = ~ 1 | category/study_type, mods= ~ qualsyst*factor(study_type), data=dat) #	  
res <- rma.mv(yi, vi, slab=dat$slab, random = ~ 1 | category, mods= ~ qualsyst*factor(study_type), data=dat)

#res <- rma(yi, vi, data=dat) #	  

### set up forest plot (with 2x2 table counts added; the 'rows' argument is
### used to specify in which rows the outcomes will be plotted)
forest(res, xlim=c(-10, 4.6), at=log(c(0.05, 0.25, 1, 4)), atransf=exp,
       ilab=cbind(sprintf("%.02f",  dat$qualsyst), dat$ni), #, dat$tneg, dat$cpos, dat$cneg),
       ilab.xpos=c(2, -5), 
	   #ilab.xpos=c(-9.5,-8,-6,-4.5), 
	   cex=cex_plot, ylim=c(-1, y_lim),
       rows=xrows, #c(3:4, 9:18, 23:27, 32:64, 69, 74:81),
	   #slab=dat$slab, 
	   #order=dat$order,
       mlab=mlabfun("RE Model for All Studies", res),
       psize=1, header="Author(s) and Year")
 
### set font expansion factor (as in forest() above) and use a bold font
op <- par(cex=0.5, font=4)
 
### add additional column headings to the plot
text(c(2), y_lim, c("QualSyst score"))
text(c(2), y_lim-1, c("score"))
text(c(-5), y_lim-1, c("n"))
 
### switch to bold italic font
par(font=4)
 
### add text for the subgroups
text(-10, sort(headings, decreasing=TRUE), pos=4, c(sprintf("Levodopa (n = %s)", sum(dat$ni[dat$category=="levodopa"])),
                               sprintf("Amantadine (n = %s)", sum(dat$ni[dat$category=="amantadine"])),
                               sprintf("Dopamine agonists (n = %s)", sum(dat$ni[dat$category=="dopamine_agonists"])), 
							   sprintf("Clozapine (n = %s)", sum(dat$ni[dat$category=="clozapine"])), 
							   sprintf("MAO-inhibitor (n = %s)", sum(dat$ni[dat$category=="mao_inhibitors"])), 
							   sprintf("Cannabis (n = %s)", sum(dat$ni[dat$category=="cannabis"])), 
							   sprintf("BOtulinum toxin (n = %s)", sum(dat$ni[dat$category=="botox"]))))
 
### set par back to the original settings
par(op)
 
### fit random-effects model in the three subgroups
res.lev <- rma(yi, vi, subset=(category=="levodopa"), data=dat)
res.ama <- rma(yi, vi, subset=(category=="amantadine"),     data=dat)
res.dag <- rma(yi, vi, subset=(category=="dopamine_agonists"),  data=dat)
res.clo <- rma(yi, vi, subset=(category=="clozapine"),  data=dat)
res.mao <- rma(yi, vi, subset=(category=="mao_inhibitors"),     data=dat)
res.cbd <- rma(yi, vi, subset=(category=="cannabis"),  data=dat)
res.bot <- rma(yi, vi, subset=(category=="botox"),  data=dat)

par(font=4)

### add summary polygons for the three subgroups
addpoly(res.lev, row= 76.5, mlab=mlabfun("RE Model for Subgroup", res.lev), cex=.5)
addpoly(res.ama, row= 70.5, mlab=mlabfun("RE Model for Subgroup", res.ama), cex=.5)
addpoly(res.clo, row= 39.5, mlab=mlabfun("RE Model for Subgroup", res.clo), cex=.5)
addpoly(res.dag, row= 30.5, mlab=mlabfun("RE Model for Subgroup", res.dag), cex=.5)
addpoly(res.mao, row= 15.5, mlab=mlabfun("RE Model for Subgroup", res.mao), cex=.5)
addpoly(res.cbd, row= 9.5, mlab=mlabfun("RE Model for Subgroup", res.cbd), cex=.5)
addpoly(res.bot, row= 1.5, mlab=mlabfun("RE Model for Subgroup", res.bot), cex=.5)
 
abline(h=0)

### fit meta-regression model to test for subgroup differences
res_sg <- rma(yi, vi, mods = ~ category, data=dat)
 
### add text for the test of subgroup differences
text(-10, -1.8, pos=4, cex=0.75, bquote(paste("Test for Subgroup Differences: ",
     "Q[M]", " = ", .(formatC(res_sg$QM, digits=2, format="f")), ", df = ", .(res_sg$p - 1),
     ", p = ", .(formatC(res_sg$QMp, digits=2, format="f")))))
	 
