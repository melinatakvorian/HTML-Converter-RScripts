
# bold_first_term_occurrence()
#
# For each row of `df`, looks across the columns in `cols` (in the
# order you give them) and, for every term in `glossary`, wraps ONLY
# the first occurrence of that term (across all selected columns)
# in <strong></strong>. Writes the result back into the same cells.
#
# - Longer/multi-word terms are matched before shorter ones, so
#   "heart attack" claims its match before "heart" does.
# - Matching uses word boundaries (\\b) so "art" won't match inside
#   "heartache".
# - Matches inside text you've already bolded (from an earlier,
#   longer term) are skipped, so you never get nested/broken tags.


escape_regex <- function(x) {
  # Escape one special character at a time using fixed = TRUE, 
  # so no regex engine ever has to parse a pattern containing these characters together.
  # IMPORTANT: backslash must be escaped first, or it would double-escape the backslashes inserted for the other characters.
  specials <- c("\\", ".", "|", "(", ")", "[", "]", "{", "}", "^", "$", "*", "+", "?")
  for (ch in specials) {
    x <- gsub(ch, paste0("\\", ch), x, fixed = TRUE)
  }
  x
}

bold_first_term_occurrence <- function(df, cols, glossary, ignore_case = TRUE) {
  
  # longest terms first, so multi-word phrases win over their substrings
  glossary <- glossary[order(-nchar(glossary))] #order glossary terms from longest to shortest
  
  for (i in seq_len(nrow(df))) { #run along each row of the data frame
    
    #create a list of the text we are investigating in df[i,] 
    #where each item is the data from one of the sections and each item's name is that section's name
      texts <- setNames(as.character(unlist(df[i, cols])), cols)
    
    # track character ranges in each column that are already inside a <strong> tag,
    #so later terms don't re-match inside them
      #I don't really understand this
      protected <- setNames(vector("list", length(cols)), cols)
      
      for (col in cols){
        protected[[col]] <- matrix(numeric(0), ncol = 2)
      }
    
    for (term in glossary) {#for each term in the glossary
      pattern <- paste0("\\b", escape_regex(term), "\\b") #create a 'tag' to assign for terms that need bolding
      found <- FALSE #set found object to FALSE
      
      for (col in cols) { #for each piece of data in 'sections'
        if (found) break #if found is true, break this code
        txt <- texts[[col]] #assign the text chunk to an object called 'txt'
        if (is.na(txt) || !nzchar(txt)) next #if txt is null or txt has no characters, then move to the next section of sections
        
        #find all the matches of term in all sections
        #store the index of each match within txt, as well as the length of the match
          m <- gregexpr(pattern, txt, ignore.case = ignore_case, perl = TRUE)[[1]] 
          if (m[1] == -1) next #if there are no mathces (m is -1), skip to the next term
          lens <- attr(m, "match.length") #takes the length of the matched words and ties it to the match's index
          prot <- protected[[col]] #the indices for the already bolded terms of this section
        
          
            for (k in seq_along(m)) { #for each match in the indices of matches (m)
              start <- m[k] #set the starting point as the m value assigned to this loop number
              end <- start + lens[k] - 1 #NOT SURE WHAT IS DONE HERE
              
              overlap <- FALSE #set overlap to false
              if (nrow(prot) > 0) { #if the number of words that are already bolded is greater than 0
                overlap <- any(start <= prot[, 2] & end >= prot[, 1]) #set overlap to TRUE if the term is already bolded
              }
              if (overlap) next #if overlap is TRUE, then go to the next match
              
              matched_text <- substr(txt, start, end) #
              new_txt <- paste0(
                substr(txt, 1, start - 1),
                "<strong>", matched_text, "</strong>",
                substr(txt, end + 1, nchar(txt))
              )
              
              # shift any already-protected ranges that came after this match
              tag_len <- nchar("<strong>") + nchar("</strong>") #
              if (nrow(prot) > 0) { #
                shift <- ifelse(prot[, 1] > end, tag_len, 0)
                prot[, 1] <- prot[, 1] + shift
                prot[, 2] <- prot[, 2] + shift
              }
              new_start <- start + nchar("<strong>")
              new_end   <- new_start + (end - start)
              prot <- rbind(prot, c(new_start, new_end))
              
              protected[[col]] <- prot
              texts[[col]] <- new_txt
              found <- TRUE
              break
            }
      }#end of loop for each piece of data in sections
    }
    
    df[i, cols] <- as.list(texts[cols])
  }
  
  df
}

#Glossary list for wildlife ----
  glossary <- c("adaptation", "AC ", "adaptive capacity", "adaptive management", "AFB", "AFS", 
                "afforestation", "anadromous", "asymptomatic reservoir", "asynchronous breeding ", "BBS", 
                "BCC", "benthic","BGEPA", "bioaccumulation", "bioclimatic variables", "bioindicators",
                "bycatch", "CEMML", "climate", "colddays", "DDT", "DOD", "DODD", "DODI", "DPS", "drought", "duration", "echolocation",
                "ecological niche model", "ecosystem engineer", "ecosystem services", "ectothermic", "el niño", "el niño-southern oscillation", 
                "ENSO", "emergent vegetation", "emissions", "endemic species", "ESA", "eutrophication",
                "evapotranspiration", "exposure", "extreme heat days", "extreme weather event", "false spring", 
                "fecundity", "flash drought", "flash flood", "frequency", "GAP", "green-up", "habitat vulnerability index", 
                "HVI", "hibernaculum", "hibernacula", "historical baseline", "hotdays", "hurricane", "hydrological drought", "ice storm", "indicator species", 
                "intensity", "interspecific brood parasitism", "IUCN", "keystone species", "la niña", "marine noise pollution", "MBTA",
                "mesic", "mesopredator", "mesopredators", "microclimate", "missionscape", "MMPA", "monsoon", "NABCI", "National Vegetation Classification System",
                "NVC", "NVCS", "natural hazard", "natural hazard exposure", "NE", "neritic", " NMFS", "NOAA",
                "Northern Atlantic Oscillation", "NAO", "ocean acidification", "OA", "other exposures", "OE", "PARC MSS", "pelagic",
                "perennial", "perennial plant", "permafrost", "permanent inundation", "phenology", "phenological", "PIF MSS", "piscivorous",
                "polyandry", "population bottleneck", "precipitation", "projection", "potential impact", "regeneration debt", "regeneration mismatch",
                "regeneration failure", "Representative Concentration Pathway",
                "RCP", "resilience", "saltwater intrusion", "scenario", "sea level decrease", "sea level increase", "SLI", "sea surface temperature",
                "SST", "sensitivity", "severity", "SFS", "SGCN", "snowpack", "SOTB TPS", "SPEI",
                "stochastic", "sp.", "spp.", "ssp.", "storm surge", " SS", "SWAP", "TED",
                "temperature-dependent sex determination", "terrestrial ecosystems", "TEVA", "torpor", "tropical cyclone", 
                "tropical storm", "typhoon", "USFWS", "USGS",
                "var.", "vulnerability", "vulnerability assessment", "VA", "weather", "wet days", 
                "white-nose syndrome", "WNS", "xeric")
#abbreviations like NE, OE, OA, SS, VA, AC, should all be checked after running to make sure random parts of words didn't get bolded


  sections_TEVA <- c("VulnSummary", "NE_Text", "OE_Text", "S_Text", "AC_Text")
  sections_FWVA <- c("ProminentTaxa", "VulnSummary", "E_Text", "S_Text", "AC_Text")
  
#unbold everything first
  if(subject == "TEVA"){
    for(col in sections_TEVA){
      df[[col]] <- stringr::str_replace_all(df[[col]], "<strong>", '')
      df[[col]] <- stringr::str_replace_all(df[[col]], "</strong>", '')
    }
  }else if (subject == "FWVA"){
    for(col in sections_FWVA){
      df[[col]] <- stringr::str_replace_all(df[[col]], "<strong>", '')
      df[[col]] <- stringr::str_replace_all(df[[col]], "</strong>", '')
    }
  }

  
#run for dataframe
  result <- bold_first_term_occurrence(df, cols = sections_TEVA, glossary = glossary)
  
  #print(result)

  
  
#create colored text for the vulnerability ----
  #the script needs to detect this text "vulnerability to short- and long-term weather changes" and find the word BEFORE it. 
  #or it needs to detect the first instance of "low", "high", "moderate", "very high" in the vulnerability summary and add the hex codes
  # <span style="color: #ff0000;">special</span>
  low_log <- str_locate(result$VulnSummary, "low <strong>vul")
  med_log <- str_locate(result$VulnSummary, "moderate <strong>vul")
  high_log <- str_locate(result$VulnSummary, "high <strong>vul")
  vhigh_log <- str_locate(result$VulnSummary, "very high <strong>vul")
  
  for(i in 1:nrow(result)){
    if(!is.na(low_log[i])){
      startval <- as.numeric(low_log[i])
      endval <- startval+2
      target <- substr(result$VulnSummary[i], startval, endval)
      before <- substr(result$VulnSummary[i], 1, startval - 1)
      after  <- substr(result$VulnSummary[i], endval + 1, nchar(result$VulnSummary[i]))
      result$VulnSummary[i] <- paste0(before, '<strong><span style="color:#8eb407;">', target, '</span></strong>', after)
      
    }else if(!is.na(vhigh_log[i])){
      startval <- as.numeric(vhigh_log[i])
      endval <- startval+8
      target <- substr(result$VulnSummary[i], startval, endval)
      before <- substr(result$VulnSummary[i], 1, startval - 1)
      after  <- substr(result$VulnSummary[i], endval + 1, nchar(result$VulnSummary[i]))
      result$VulnSummary[i] <- paste0(before, '<strong><span style="color:#d42004;">', target, '</span></strong>', after)
      
    }else if(!is.na(med_log[i])){
      startval <- as.numeric(med_log[i])
      endval <- startval+7
      target <- substr(result$VulnSummary[i], startval, endval)
      before <- substr(result$VulnSummary[i], 1, startval - 1)
      after  <- substr(result$VulnSummary[i], endval + 1, nchar(result$VulnSummary[i]))
      result$VulnSummary[i] <- paste0(before, '<strong><span style="color:#BCC208;">', target, '</span></strong>', after)
      
      
    }else if(!is.na(high_log[i])){
      startval <- as.numeric(high_log[i])
      endval <- startval+3
      target <- substr(result$VulnSummary[i], startval, endval)
      before <- substr(result$VulnSummary[i], 1, startval - 1)
      after  <- substr(result$VulnSummary[i], endval + 1, nchar(result$VulnSummary[i]))
      result$VulnSummary[i] <- paste0(before, '<strong><span style="color:#f49e0b;">', target, '</span></strong>', after)
      
    }else{
      next}
  }

#EXPORT ----
  ##export excel to 3ViewerPackages folder ----
  out_dir <- paste0(input_umbrella, input_installation_folder, "/3ViewerPackages/HTML_excels") 
  
  # ******** NOTE THAT THE FOLDER STRUCTURE MUST MATCH WHAT IS ABOVE ^^^ EXACTLY.  **********
  # CHANGE out_dir AS NEEDED IF THERE ARE ANY DIFFERENCES IN THE LOCATION YOU WANT TO SAVE TO.
  
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  
  output_filename <- paste0(project_name, "_HTML_formatted.xlsx")
  shortcut_location <- file.path(input_dir, output_filename) #save the path to the future shortcut
  
  write_xlsx(result, shortcut_location) #create file and save to folder
  message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# ---------------------------------------------------------------
# Example usage ----
df <- data.frame(
  colA = c("Patient had a heart attack.", "No cardiac issues noted."),
  colB = c("The heart attack occurred at home.", "Follow-up for heart attack risk."),
  stringsAsFactors = FALSE
)

glossary <- c("heart attack", "heart")

result <- bold_first_term_occurrence(df, cols = c("colA", "colB"), glossary = glossary)
print(result)
#
# Expected for row 1:
#  - "heart attack" (as a phrase) gets its first occurrence bolded in
#    colA: "Patient had a <strong>heart attack</strong>."
#  - "heart" (as its own glossary term) is NOT re-bolded inside that
#    same "heart attack" span (it's protected), so its own first
#    occurrence is found next, in colB:
#    "The <strong>heart</strong> attack occurred at home."
#  Each glossary term is tracked independently -- "heart" and "heart
#  attack" both get exactly one bolded occurrence each, and neither
#  steals territory already claimed by the other.