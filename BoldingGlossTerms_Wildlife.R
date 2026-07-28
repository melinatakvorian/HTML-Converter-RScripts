#BOLDING GLOSSARY TERMS
#TESTING WITH CREECH AFB
#07-28-2026
#Author: Melina Takvorian

#This script will bold the first instance of any terms in a given section that appear in the list of glossary terms

#Create list of glossary terms ----
terms <- c("weather", "vulnerability", "natural hazards")

#Identify sections that need to be checked ----
sections <- c("VulnSummary", "NE_Text", "OE_Text", "S_Text", "AC_Text")
sections_idx <- c(24, 25, 28, 31, 34)




#search and replace the instances of these terms ----
  #first, go through the first section of the document and store the instances of the 
low_log <- str_locate(df$VulnSummary, "low <strong")

for(i in 1:nrow(df)){
  if(!is.na(low_log[i])){
    startval <- as.numeric(low_log[i])
    endval <- startval+2
    target <- substr(df$VulnSummary[i], startval, endval)
    before <- substr(df$VulnSummary[i], 1, startval - 1)
    after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
    df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#8eb407;">', target, '</span></strong>', after)
    
  }else{
    next}
}
