## Functions Library
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2026.09.21


# ----- * Word->HTML function ----
# takes Word document (input) and turns it into HTML file (output)

convert_docx_to_html_full <- function(docx_file) {
  html_file <- tempfile(fileext = ".html")
  
  pandoc::pandoc_convert(
    file = docx_file,
    output = html_file,
    from = "docx",
    to = "html",
    standalone = TRUE,
    args = c("--wrap=none")
  )
  
  xml2::read_html(html_file)
}



# ----- * HTML->pieces function ----
# reads HTML file (input) and separate sections for building table later 
parse_html_sections <- function(html_doc) {
  headings <- rvest::html_nodes(html_doc, "h1") #identify headings
  sections <- vector("list", length(headings)) #create list of headings (sections)
  
  for (i in seq_along(headings)) { #for each section, concatenate all the info that belongs to it (across docs)
    start_node <- headings[[i]]
    
    end_node <- if (i < length(headings)) headings[[i + 1]] else NULL
    siblings <- xml2::xml_find_all(start_node, "following-sibling::*")
    if (!is.null(end_node)) {
      idx <- which(vapply(siblings, identical, logical(1), y = end_node))
      if (length(idx) == 0) idx <- length(siblings) + 1
      siblings <- siblings[seq_len(idx - 1)]
    }
    
    
    # Insert a space between concatenated HTML nodes
    content_html <- paste(as.character(siblings), collapse = " ")
    sections[[i]] <- list(title = xml_text(start_node), content = content_html)
  }
  
  names(sections) <- vapply(sections, `[[`, "", "title")
  lapply(sections, `[[`, "content")
}

# ----- * removing spaces after headings function -----
#if results[i] ends with " ", remove it
remove_end_blanks <- function(result_list){
  
  for(i in 1:length(result_list)){
    templist <- result_list[[i]]
    
    for(heading in 1:length(templist)){
      if(endsWith(names(templist)[heading], " ")){
        
        headingWithSpace <- names(templist)[heading] #save heading to local object
        print(headingWithSpace)
        
        endstring <- stringr::str_length(headingWithSpace) #find length of heading's string
        
        endstring <- as.numeric(endstring)-1
        
        headingNoSpace <- substr(headingWithSpace, 1, endstring) #remove space from end and save
        #print(headingNoSpace)
        
        names(result_list[[i]])[heading] <- headingNoSpace
        print(names(result_list[[i]][heading]))
      }else next
    }
  }
  return(result_list)
}

# ----- * remove '\r\n' from heading names -----
#if results[i] includes '\r\n', remove it
remove_accidental_return <- function(result_list){
  
  for(i in 1:length(result_list)){
    templist <- result_list[[i]]
    
    for(heading in 1:length(templist)){
      if(stringr::str_detect(names(templist)[heading], "\\r\\n")){
        
        #replace "\r\n" with nothing
        headingWithProblem <- names(templist)[heading] #save heading to local object
        
        newHeading <- stringr::str_replace_all(headingWithProblem, "\\r\\n", " ")
        
        names(result_list[[i]])[heading] <- newHeading
        print(names(result_list[[i]][heading]))
      }else next
    }
  }
  return(result_list)
}


# ----- * replace all except the last instance of a substring -----
replace_all_except_last <- function(s, from, to) {
  # Find the last occurrence of `from`
  matches <- gregexpr(from, s, fixed = TRUE)[[1]]
  
  # No occurrences — return as-is
  if (matches[1] == -1){return(s)}
  
  last_pos <- tail(matches, 1)
  last_len <- attr(matches, "match.length") |> tail(1)
  
  # Split into before and after (inclusive of) the last match
  before <- substr(s, 1, last_pos - 1)
  after  <- substr(s, last_pos, nchar(s))
  
  # Replace all occurrences in the prefix, leave the tail unchanged
  paste0(gsub(from, to, before, fixed = TRUE), after)
}

# * remove paragraph notation ----
p_be_gone <-  function(df, columns){
  for(col in columns){
    if (!col %in% colnames(df)) {
      warning(paste("Column not found, skipping:", col))
      next
    }
    #remove paragraph notation
    df[[col]] <- stringr::str_replace_all(df[[col]], "<p>", '')
    df[[col]] <- stringr::str_replace_all(df[[col]], "</p>", '')
    
  }
  return(df)
}

# * hanging indents for references ----
ref_hanging_indents <- function(df, report_type){
  if(report_type == "TEVA"){
    for(i in 1:nrow(df)){
      #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
      temp_string <- df$`References and Credits`[i]
      temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
      df$`References and Credits`[i] <- temp_string1 
    }
    return(df)
  }else if(report_type == "FWVA"){
    # for(i in 1:nrow(df)){
    #   df$References[i] <- stringr::str_replace_all(df$References[i], "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
    # }
    for(i in 1:nrow(df)){
      #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
      temp_string <- df$`References`[i]
      temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
      df$`References`[i] <- temp_string1 #change to temp_string2 if you are adding the line breaks
    }
    return(df)
  }
}
# * update U.S. to US ----
#THE ONLY USE-CASE THIS DOES NOT HANDLE IS WHEN U.S. IS THE LAST WORD OF THE LAST SENTENCE OF THE STRING.
#THIS FUNCTION WILL MAKE THAT U.S. -> US, WHERE THERE IS NO PERIOD TO END THE SENTENCE

update_US <- function(df, report_type, installation_type){
  if(installation_type == "Air Force" && report_type == "TEVA"){
    
    cols_to_search <- c(24,25,28,31,34) #the indices of VulnSummary, NE_Text, OE_Text, S_Text, AC_Text
    
    for(col in cols_to_search){ 
      df[[col]] <- gsub("U\\.S\\. ([A-Z])", "US. \\1", df[[col]])  # detect capital letters indicating a new sentence
      df[[col]] <- gsub("U\\.S\\.<sup", "US.<sup", df[[col]]) #detect superscripted numbers indicating a new sentence
      df[[col]] <- gsub("U\\.S\\.", "US", df[[col]])                # everything else
    }
    
    return(df)
    
  }else if(installation_type == "Air Force" && report_type == "FWVA"){
    cols_to_search <- c(12, 13, 15, 17) #the indices of VulnSummary, E_Text, S_Text, AC_Text
    
    for(col in cols_to_search){ 
      df[[col]] <- gsub("U\\.S\\. ([A-Z])", "US. \\1", df[[col]])  # detect capital letters indicating a new sentence
      df[[col]] <- gsub("U\\.S\\.<sup", "US.<sup", df[[col]]) #detect superscripted numbers indicating a new sentence
      df[[col]] <- gsub("U\\.S\\.", "US", df[[col]])                # everything else
    }
    
    return(df)
    
  }
}

# * assign Hex codes and Numeric values to columns that need it -----
hex_codes <- function(df, report_type){
  if(report_type == "TEVA"){
    ##TEVAs
    #repeat this for VulnerabilityResult, Confidence, NE_Level, OE_Level, S_Level, AC_Level
    
    #VulnNum
    df <- df %>% 
      mutate('VulnNum' = case_when(
        VulnerabilityResult == "VERY HIGH" ~ 4,
        VulnerabilityResult == "HIGH" ~ 3,
        VulnerabilityResult == "MODERATE" ~ 2,
        VulnerabilityResult == "LOW" ~ 1,
        TRUE ~ 1
      )) %>% relocate('VulnNum', .after = VulnerabilityResult)
    
    
    
    #VulnColor
    df <- df %>% 
      mutate(VulnColor = case_when(
        VulnerabilityResult == "VERY HIGH" ~ "#d42004",
        VulnerabilityResult == "HIGH" ~ "#f49e0b",
        VulnerabilityResult == "MODERATE" ~ "#f2e750",
        VulnerabilityResult == "LOW" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(VulnColor, .after = VulnerabilityResult)
    
    #Confidence
    df <- df %>% 
      mutate('ConfNum' = case_when(
        Confidence == "High" ~ 3,
        Confidence == "Moderate" ~ 2,
        Confidence == "Low" ~ 1,
        TRUE ~ 1
      )) %>% relocate('ConfNum', .after = Confidence)
    
    #NE_Level
    df <- df %>% 
      mutate(NE_Color = case_when(
        NE_Level == "High" ~ "#f49e0b",
        NE_Level == "Moderate" ~ "#f2e750",
        NE_Level == "Low" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(NE_Color, .after = NE_Level)
    
    #OT_Level
    df <- df %>% 
      mutate(OE_Color = case_when(
        OE_Level == "High" ~ "#f49e0b",
        OE_Level == "Moderate" ~ "#f2e750",
        OE_Level == "Low" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(OE_Color, .after = OE_Level)
    
    #S_Level
    df <- df %>% 
      mutate(S_Color = case_when(
        S_Level == "High" ~ "#f49e0b",
        S_Level == "Moderate" ~ "#f2e750",
        S_Level == "Low" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(S_Color, .after = S_Level)
    
    #AC_Text
    #this one is different from the rest!!
    df <- df %>% 
      mutate(AC_Color = case_when(
        AC_Level == "High" ~ "#b2e109",
        AC_Level == "Moderate" ~ "#f2e750",
        AC_Level == "Low" ~ "#f49e0b",
        TRUE ~ "none"
      )) %>% relocate(AC_Color, .after = AC_Level)
  }else if(report_type == "FWVA"){
    ##FWVAs
    #repeat this for VulnerabilityResult, E_Level, S_Level, AC_Level
    
    #VulnNum
    df <- df %>% 
      mutate('VulnNum' = case_when(
        VulnerabilityResult == "VERY HIGH" ~ 4,
        VulnerabilityResult == "HIGH" ~ 3,
        VulnerabilityResult == "MODERATE" ~ 2,
        VulnerabilityResult == "LOW" ~ 1,
        TRUE ~ 1
      )) %>% relocate('VulnNum', .after = VulnerabilityResult)
    
    #VulnColor
    df <- df %>% 
      mutate(VulnColor = case_when(
        VulnerabilityResult == "VERY HIGH" ~ "#d42004",
        VulnerabilityResult == "HIGH" ~ "#f49e0b",
        VulnerabilityResult == "MODERATE" ~ "#f2e750",
        VulnerabilityResult == "LOW" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(VulnColor, .after = 'VulnNum')
    
    #E_Level
    df <- df %>% 
      mutate(E_Color = case_when(
        E_Level == "High" ~ "#f49e0b",
        E_Level == "Moderate" ~ "#f2e750",
        E_Level == "Low" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(E_Color, .after = E_Level)
    
    
    #S_Level
    df <- df %>% 
      mutate(S_Color = case_when(
        S_Level == "High" ~ "#f49e0b",
        S_Level == "Moderate" ~ "#f2e750",
        S_Level == "Low" ~ "#b2e109",
        TRUE ~ "none"
      )) %>% relocate(S_Color, .after = S_Level)
    
    
    #AC_Level
    df <- df %>% 
      mutate(AC_Color = case_when(
        AC_Level == "High" ~ "#b2e109",
        AC_Level == "Moderate" ~ "#f2e750",
        AC_Level == "Low" ~ "#f49e0b",
        TRUE ~ "none"
      )) %>% relocate(AC_Color, .after = AC_Level)
  }
}


# * add Habitat_Icon columns ----
habitat_icons <- function(df){
  df[,'FirstHabitatIcon'] <- ""
  df[,'SecondHabitatIcon'] <- ""
  df[,'ThirdHabitatIcon'] <- ""
  df[,'FourthHabitatIcon'] <- ""
  df <- df %>% 
    relocate('FirstHabitatIcon', .after = `FirstHabitat`) %>% 
    relocate('SecondHabitatIcon', .after = `SecondHabitat`) %>% 
    relocate('ThirdHabitatIcon', .after = `ThirdHabitat`) %>%
    relocate('FourthHabitatIcon', .after = `FourthHabitat`)
}

# * color vulnerability text ----
  # color_vuln_text <- function(df){
  #   #create colored text for the vulnerability 
  #   #the script needs to detect this text "vulnerability to short- and long-term weather changes" and find the word BEFORE it. 
  #   #or it needs to detect the first instance of "low", "high", "moderate", "very high" in the vulnerability summary and add the hex codes
  #   # <span style="color: #ff0000;">special</span>
  #   
  #   low_log <- str_locate(df$VulnSummary, "low <strong")
  #   med_log <- str_locate(df$VulnSummary, "moderate <strong")
  #   high_log <- str_locate(df$VulnSummary, "high <strong")
  #   vhigh_log <- str_locate(df$VulnSummary, "very high <strong")
  #   
  #   for(i in 1:nrow(df)){
  #     if(!is.na(low_log[i])){
  #       startval <- as.numeric(low_log[i])
  #       endval <- startval+2
  #       target <- substr(df$VulnSummary[i], startval, endval)
  #       before <- substr(df$VulnSummary[i], 1, startval - 1)
  #       after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
  #       df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#8eb407;">', target, '</span></strong>', after)
  #       
  #     }else if(!is.na(vhigh_log[i])){
  #       startval <- as.numeric(vhigh_log[i])
  #       endval <- startval+8
  #       target <- substr(df$VulnSummary[i], startval, endval)
  #       before <- substr(df$VulnSummary[i], 1, startval - 1)
  #       after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
  #       df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#d42004;">', target, '</span></strong>', after)
  #       
  #     }else if(!is.na(med_log[i])){
  #       startval <- as.numeric(med_log[i])
  #       endval <- startval+7
  #       target <- substr(df$VulnSummary[i], startval, endval)
  #       before <- substr(df$VulnSummary[i], 1, startval - 1)
  #       after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
  #       df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#BCC208;">', target, '</span></strong>', after)
  #       
  #       
  #     }else if(!is.na(high_log[i])){
  #       startval <- as.numeric(high_log[i])
  #       endval <- startval+3
  #       target <- substr(df$VulnSummary[i], startval, endval)
  #       before <- substr(df$VulnSummary[i], 1, startval - 1)
  #       after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
  #       df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#f49e0b;">', target, '</span></strong>', after)
  #       
  #     }else next
  #   }
  #   return(df)
  # }
  
color_vuln_text <- function(df){
  ##doing it without bolding ----
    #the script needs to detect this text "vulnerability to short- and long-term weather changes" and find the word BEFORE it.
    #or it needs to detect the first instance of "low", "high", "moderate", "very high" in the vulnerability summary and add the hex codes
    # <span style="color: #ff0000;">special</span>
    low_log <- str_locate(df$VulnSummary, "low vul")
    med_log <- str_locate(df$VulnSummary, "moderate vul")
    high_log <- str_locate(df$VulnSummary, "high vul")
    vhigh_log <- str_locate(df$VulnSummary, "very high vul")

    for(i in 1:nrow(df)){
      if(!is.na(low_log[i])){
        startval <- as.numeric(low_log[i])
        endval <- startval+2
        target <- substr(df$VulnSummary[i], startval, endval)
        before <- substr(df$VulnSummary[i], 1, startval - 1)
        after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
        df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#8eb407;">', target, '</span></strong>', after)

      }else if(!is.na(vhigh_log[i])){
        startval <- as.numeric(vhigh_log[i])
        endval <- startval+8
        target <- substr(df$VulnSummary[i], startval, endval)
        before <- substr(df$VulnSummary[i], 1, startval - 1)
        after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
        df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#d42004;">', target, '</span></strong>', after)

      }else if(!is.na(med_log[i])){
        startval <- as.numeric(med_log[i])
        endval <- startval+7
        target <- substr(df$VulnSummary[i], startval, endval)
        before <- substr(df$VulnSummary[i], 1, startval - 1)
        after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
        df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#BCC208;">', target, '</span></strong>', after)


      }else if(!is.na(high_log[i])){
        startval <- as.numeric(high_log[i])
        endval <- startval+3
        target <- substr(df$VulnSummary[i], startval, endval)
        before <- substr(df$VulnSummary[i], 1, startval - 1)
        after  <- substr(df$VulnSummary[i], endval + 1, nchar(df$VulnSummary[i]))
        df$VulnSummary[i] <- paste0(before, '<strong><span style="color:#f49e0b;">', target, '</span></strong>', after)

      }else next
    }
    return(df)
}
  
    
