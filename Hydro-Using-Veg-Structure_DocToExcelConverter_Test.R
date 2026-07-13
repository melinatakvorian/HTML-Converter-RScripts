# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2026.04.28


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

packages <- c("pandoc","xml2","rvest","writexl", "readxl","dplyr")
# Install packages not yet installed
installed_packages <- packages %in% rownames(installed.packages())
if (any(installed_packages == FALSE)) {
  install.packages(packages[!installed_packages]) #error here
}

# load packages
invisible(lapply(packages, library, character.only = TRUE))

## Create paths for storing files ----

#####CHANGE AS DIRECTED BELOW --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

#PAY ATTENTION TO THE DIRECTION OF THE SLASHES. THEY HAVE TO BE CHANGED TO FORWARD SLASHES, AS SHOWN BELOW
#the broad folder structure
 
# ----TEXT FOR YOU TO CHANGE-----------
  # Select which installation folder you're working in
  input_installation_folder <- "Cold Bay LRRS"

  # Write if working on AF (AIR FORCE) or Navy (NAVY):
  # inst_sheet = "NAVY"
  inst_sheet = "AIR FORCE"

  # If Navy, select which region
  # navy_region = "Southeast Region"
  
  # Select which analysis you're doing (shouldn't need to change)
  input_SME_folder <- "/Hydrology/Word to HTML" 
  
  #the final file name will start with this and will get the date added
  subject <- "Hydro"
  project_name <- paste0(subject, "_", input_installation_folder)
  
  # this will select which base to select your data from
  ifelse(inst_sheet == "AIR FORCE",
         input_umbrella <- "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/",
         input_umbrella <- paste0("N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/", navy_region, "/"))


#####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

  input_dir <-  paste0(input_umbrella, input_installation_folder, input_SME_folder)
  current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"
  installation_info <- readxl::read_xlsx("Installation_IDs.xlsx", sheet = inst_sheet)
  
#ERROR CATCH 1: open file ----

filenames <- list.files(input_dir) #create list of files in the folder
openfiles <- list()

for(file in 1:length(filenames)){ #check that there are no open files
  if(startsWith(filenames[file], "~")){
    openfiles[length(openfiles)+1] <- filenames[file]
  }else next
} 

if(!length(openfiles)==0){
  stop("The following document(s) is open on a computer. This script cannot run unless all files are closed.\n 
           Open files are listed below: \n", openfiles)
}


# ----- * Word->HTML function ----
# takes Word document (input) and turns it into HTML file (output)
convert_docx_to_html_full <- function(docx_file, filepath) {
  #html_file <- tempfile(fileext = ".html")
  html_file <- paste0(filepath, "/output1.html")
  
  pandoc::pandoc_convert(
    file = docx_file,
    output = html_file,
    from = "docx",
    to = "html",
    standalone = TRUE
  )
  
  xml2::read_html(html_file)
}


# ----- * HTML->pieces function ----
#reads HTML file (input) and separate sections for building table later
parse_html_sections_hist <- function(html_doc, section_indices) {
  #identify all headings
  headings <- rvest::html_nodes(html_doc, "h1") #identify headings
  sections <- vector("list", length(section_indices)) #create list of headings (sections)
  
  for (i in seq_along(section_indices)) { # Iterate over specified sections
    print(paste("Parsing section:", i))
    
    start_node <- headings[[section_indices[i]]]
    
    end_node <- if (i < length(section_indices)) headings[[i + 1]] else NULL
    
    #print(headings[section_indices[i + 1]])
    
    siblings <- xml2::xml_find_all(start_node, "following-sibling::*") #the entire rest of the doc?
    if (!is.null(end_node)) {
      idx <- which(vapply(siblings, identical, logical(1), y = end_node))
      if (length(idx) == 0) idx <- length(siblings) + 1
      siblings <- siblings[seq_len(idx - 1)]
    }
    
    # Insert a space between concatenated HTML nodes
    content_html <- paste(as.character(siblings), collapse = " ")
    sections[[i]] <- content_html
    print(content_html)
  }
  
  # Assign section titles as names to the list elements
  names(sections) <- sapply(headings[section_indices], xml_text)
  sections
  
}


parse_html_sections_disr <- function(html_doc, section_indices) {
  #identify all headings
  headings <- rvest::html_nodes(html_doc, "h1") #identify headings
  sections <- vector("list", length(section_indices)) #create list of headings (sections)
  
  for (i in seq_along(section_indices)) { # Iterate over specified sections
    print(paste("Parsing section:", i))
    
    start_node <- headings[[section_indices[i]]]
    
    end_node <- if (i <= length(section_indices)) headings[[section_indices[i]+1]] else NULL
    #end_node <- if (i < length(section_indices)) headings[[i + 1]] else NULL
    #print(headings[section_indices[i + 1]])
    
    siblings <- xml2::xml_find_all(start_node, "following-sibling::*")
    if (!is.null(end_node)) {
      idx <- which(vapply(siblings, identical, logical(1), y = end_node))
      if (length(idx) == 0) idx <- length(siblings) + 1
      siblings <- siblings[seq_len(idx - 1)]
    }
    
    # Insert a space between concatenated HTML nodes
    content_html <- paste(as.character(siblings), collapse = " ")
    sections[[i]] <- content_html
    print(content_html)
  }
  
  # Assign section titles as names to the list elements
  names(sections) <- sapply(headings[section_indices], xml_text)
  sections
  
}
# ----- * removing spaces after headings function -----
#if sections[i] ends with " ", remove it
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

# ----- * replace the last instance of a substring -----
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

# RUN ----

#initialize objects for storing file info
docx_files <- list.files(input_dir, pattern = "\\.docx$", full.names = TRUE) #pull list of all files in folder
results_hist <- list()
results_disr<- list()

#for each file, convert it to HTML, Identify its sections, delete empty headers, add to a results mega-list
for (file in docx_files) { 
  html_doc <- convert_docx_to_html_full(file, input_dir)
  
  #identify all headings
  headings <- rvest::html_nodes(html_doc, "h1")
  
  #create a list of all heading names
  nlist <- c()
  for(i in seq_along(headings)){
    temp <- xml_attr(headings[[i]], "id")
    nlist[length(nlist)+1] <- temp
  }
  
  last <- as.numeric(length(nlist))
  
  ##### MA UPDATED THIS WITH NEW INDICES GIVEN REMOVAL OF SITENAME COLUMN
  # Define indices for histclimatic and disruptions sections
    hist_indices <- c(1:5, last) # histclimatic sections
    disr_indices <- c(1, 6:(last-1)) # disruptions sections
  
  
  #Create hist table list
  sections_hist <- parse_html_sections_hist(html_doc, hist_indices)
  sections_hist <- sections_hist[names(sections_hist) != ""] #remove accidental headers
  results_hist[[basename(file)]] <- sections_hist #should be a list of headings and its text
  
  #Create disr table list
  sections_disr <- parse_html_sections_disr(html_doc, disr_indices)
  sections_disr <- sections_disr[names(sections_disr) != ""] #remove accidental headers
  results_disr[[basename(file)]] <- sections_disr #should be a list of headings and its text
}

#remove blank spaces after headings that could cause additional headers accidentally
results_hist <- remove_end_blanks(results_hist)
results_disr <- remove_end_blanks(results_disr)

results_disr <- remove_accidental_return(results_disr)
results_hist <- remove_accidental_return(results_hist)


#unfold the results list to be able to create a dataframe *************** MA - changed the headings in all_headings_disr
all_headings_hist <- unique(unlist(lapply(results_hist, names)))
all_headings_disr <- unique(unlist(lapply(results_disr, names)))
# all_headings_disr <- "SITEID"
# all_headings_disr <- append(all_headings_disr, unique(unlist(lapply(results_disr, names))))

# Create dataframe and input HTML in proper sections ----

##organize historic scenario ----
df_hist <- data.frame(matrix(NA_character_, length(results_hist), length(all_headings_hist)),
                     stringsAsFactors = FALSE)
colnames(df_hist) <- all_headings_hist
rownames(df_hist) <- names(results_hist)

for (i in seq_along(results_hist)) {
  for (col in all_headings_hist) {
    if (col %in% names(results_hist[[i]])) {
      df_hist[i, col] <- results_hist[[i]][[col]]
    }else{df_hist[i, col] <- NA} 
  }
}

##### Testing new column for Sitename   ************* ADDED BY MA TO ADD SITENAME BACK - add code to remove <>
df_hist$SITENAME <- NA
# move it to the front
total_cols <- ncol(df_hist)
df_hist <- df_hist[, c(total_cols, 1:(total_cols - 1))]



##identify disruption scenarios----

#find the indices within the list that are new occurrences of 'New_Scenario'
num_files <- as.list(c(1:as.numeric(length(results_disr)))) #initialize list

#create mini lists for each instance of new disruption scenario
for(file in seq_along(results_disr)){
  disr_names <- names(results_disr[[file]]) #create list of headings from each file
  indices <- c()
  
  for(i in seq_along(disr_names)){ 
    if(disr_names[i] =="New_Scenario"){  #if the heading at this index is a new disr group
      indices[length(indices)+1] <- i} #save the index to the end of the indices list
  }
  num_files[[file]] <- indices #append the indices of new disr group to this number file in the folder
}

#use the indices to create smaller lists as keys to sections of disr Groups in the document
#initialize objects
total_rows <- 0
mylist <- vector("list", length(results_disr))

#create mini lists, assign data to them
for(file in seq_along(results_disr)){ #for each heading
  
  for(i in seq_along(num_files[[file]])){ #in each file
    finish <- i+1 
    
    if(finish <= length(num_files[[file]])){ #if we aren't past the last file in the folder,
      secondtolast <- num_files[[file]][[finish]] #get the instance of the next New disr heading
      secondtolast <- secondtolast-1 #we want to stop BEFORE we get to the next section
      
      num_pair <- c(num_files[[file]][[i]]:secondtolast) #create range from one to the next
      
      total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
      
      mylist[[file]][[length(mylist[[file]])+1]] <- num_pair #create a nested list with each index within a disruption scenario section
      
    }else{ #case for the last instance of new disruption scenario
      num_pair <- c(num_files[[file]][[i]]:length(results_disr[[file]]))
      total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
      
      mylist[[file]][[length(mylist[[file]])+1]] <- num_pair
    }
  }
}


#create a df where each row is one of these lists. 
df_disr <- data.frame(matrix(NA_character_, nrow=length(total_rows), ncol=length(unique(all_headings_disr))),
                     stringsAsFactors = FALSE)
colnames(df_disr) <- unique(all_headings_disr)
rownum <- 1

##### Testing new column for Sitename   ************* ADDED BY MA TO ADD SITENAME BACK
df_disr$SITENAME <- NA
# move it to the front
total_cols <- ncol(df_disr)
df_disr <- df_disr[, c(total_cols, 1:(total_cols - 1))]

for(file in seq_along(results_disr)){
  for(a in seq_along(mylist[[file]])){
    # Extract the current list of indices from mylist
    templist <- mylist[[file]][[a]]
    
    # Populate the first few columns with results_hist data (assuming it applies to all rows for this file)
    df_disr[rownum, 2] <- results_hist[[file]][[1]] #########MA CHANGED THIS, NEEDED TO ADD TO SECOND COL NOT FIRST. GOT RID OF SECOND SINCE NO SITENAME TO POPULATE
    # df_disr[rownum, 1] <- results_hist[[file]][[2]] #COMMENTED OUT SINCE SITENAME IS NO LONGER INCLUDED IN DOC
    
    n_col <- 3 # Start filling from the 4th column
    
    # Extract elements from results_disr based on the indices in templist
    for(b in seq_along(templist)){
      df_disr[rownum, n_col] <- results_disr[[file]][[templist[[b]]]]
      n_col <- n_col + 1
    }
    
    # Move to the next row for the dataframe
    rownum <- rownum + 1
  }
}

#TRANSPOSE DATATABLES ----

#create initial dataframe structure
  scenario <- c("Historical", "Moderate Disruption", "Moderate Disruption", "High Disruption", "High Disruption")
  period <- c("Historical", "Near Term", "Far Term", "Near Term", "Far Term")
  
  colnames <- c("SITENAME", "SITEID", "Scenario", "Period","SPEI_Text", "Installation_Summary", "Dry_Distribution_Text", 
                "Wet_Distribution_Text", "Dry_Duration_Severity_Text", "Wet_Duration_Severity_Text", 
                "References")
  
  ##Create mini-tables based on column names----
  all_disr_names <- colnames(df_disr)
  near_term <- all_disr_names[stringr::str_starts(all_disr_names,"Period: Near Term")]
  far_term <- all_disr_names[stringr::str_starts(all_disr_names,"Period: Far Term")]
  
  df_near_term <- df_disr %>% 
    select(all_of(near_term))
  
  df_far_term <- df_disr %>% 
    select(all_of(far_term))
  
  leftovers <- df_disr %>% 
    select(-c(all_of(near_term), all_of(far_term)))
  
  #create dataframe
    all_scenarios_df <- matrix(nrow = 5, ncol = length(colnames))
    all_scenarios_df <- as.data.frame(all_scenarios_df)
    colnames(all_scenarios_df) <- colnames
    all_scenarios_df$Scenario <- scenario
    all_scenarios_df$Period <- period
  
    all_scenarios_df2 <- all_scenarios_df
  
  ###start transposing data starting AFTER historical row----
    inst_summ <- as.numeric(which(colnames(df_hist) == "Installation_Summary"))
    
    #Fill SITENAME, SITEID, Installation Summary ***************************CHANGE THIS CODE******
    all_scenarios_df2$Installation_Summary <- df_hist[1,inst_summ]
    all_scenarios_df2$SITENAME <- df_hist$SITENAME
    all_scenarios_df2$SITEID <- df_hist$SITEID
    
    #Fill row 1 with historical data
    all_scenarios_df2$SPEI_Text[1] <- df_hist$`Period: Historical, SPEI_Text`[1]
    all_scenarios_df2$References[1] <- df_hist$References[1]
    
    
    # Fill rows 2,4 for NEAR TERM scenarios
    all_scenarios_df2$SPEI_Text[[2]] <- df_near_term$`Period: Near Term, SPEI_Text`[1]
    all_scenarios_df2$SPEI_Text[[4]] <- df_near_term$`Period: Near Term, SPEI_Text`[2]

    
    # Fill rows 3,5 for FAR TERM scenarios
    all_scenarios_df2$SPEI_Text[[3]] <- df_far_term$`Period: Far Term, SPEI_Text`[1]
    all_scenarios_df2$SPEI_Text[[5]] <- df_far_term$`Period: Far Term, SPEI_Text`[2]

    
    #Fill in the other rows based on disruption
    all_scenarios_df2[c(4,5), 7:10] <- leftovers[2, 5:8] #high disruption
    all_scenarios_df2[c(2,3), 7:10] <- leftovers[1, 5:8] #moderate disruption
    
#add BLANK numeric columns----
  
  cols <- as.numeric(ncol(all_scenarios_df2))+1
  cols_w_nos <- cols + 7
  all_scenarios_df2[,c(cols:cols_w_nos)] <- ""
  
  #numeric column names
  new_cols <- c("Minimum_SPEI", "Maximum_SPEI", 
                "Dry_Variability", "Wet_Variability", "Dry_Events", "Wet_Events", 
                "Dry_Change", "Wet_Change", "")
  
  #assign names
  colnames(all_scenarios_df2)[cols:cols_w_nos] <- new_cols #this one errors, don't worry about it
  
  #move columns to where Anthony wants them
  all_scenarios_df2 <- all_scenarios_df2 %>% 
    relocate(all_of(cols:cols_w_nos), .before = "Installation_Summary")
    
  all_scenarios_df3 <- all_scenarios_df2
  
#Formatting indents and line breaks ----
  ##add hanging indent to references section ----
    for(i in 1:nrow(all_scenarios_df3)){
      if(is.na(all_scenarios_df3$References[i])) next #skip NA rows
      
      all_scenarios_df3$References[i]
      
      #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
      temp_string <- all_scenarios_df3$References[i]
      
      #temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style="padding-left:15px;text-indent:-15px;">')
      temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
      temp_string2 <- replace_all_except_last(temp_string1, "</p>", "</p> <br>")
      all_scenarios_df3$References[i] <- temp_string2
    }
  
  ##add line breaks to '_Text' columns ----
  
    #these are the sections that need the line breaks after each paragraph:
      #"SPEI_Text", "Dry_Distribution_Text", "Wet_Distribution_Text", 
      #"Dry_Duration_Severity_Text", "Wet_Duration_Severity_Text"
    
    numbblocks <- c(5, 15:18) #corresponds to the columns with text that we need broken up

      for(a in 1:length(numbblocks)){
        col_num <- numbblocks[[a]]
        for(b in 1:nrow(all_scenarios_df3)){
          if(is.na(all_scenarios_df3[[col_num]][b])) next
          
          #replace each </p> to </p> <br>
          temp_string <- all_scenarios_df3[[col_num]][b]
          temp_string1 <- replace_all_except_last(temp_string, "</p>", "</p> <br>")
          all_scenarios_df3[[col_num]][b] <- temp_string1
        }
      }
  
    ##add indent at the beginning of each non-bulleted paragraph ----
      for(a in 1:length(numbblocks)){
        col_num <- numbblocks[[a]]
        for(b in 1:nrow(all_scenarios_df3)){
          
          if(is.na(all_scenarios_df3[[col_num]][b])) next
          
          #replace each <p> to <p style=text-indent:-15px;>
          temp_string <- all_scenarios_df3[[col_num]][b]
          temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=text-indent:15px;>')
          all_scenarios_df3[[col_num]][b] <- temp_string1
        }
      }
    
    ##add blank line after each bulleted paragraph ----
      for(i in 1:nrow(all_scenarios_df3)){
        if(is.na(all_scenarios_df3$Installation_Summary[i])) next
        
        #replace each </p></li> to </p></li><br>
        temp_string <- all_scenarios_df3$Installation_Summary[i]
        temp_string1 <- replace_all_except_last(temp_string, "</p></li>", "</p></li><br>")
        all_scenarios_df3$Installation_Summary[i] <- temp_string1
      }
  
    ##add blank line after the subheading in Installation Summary ----
      for(i in 1:nrow(all_scenarios_df3)){
        all_scenarios_df3$Installation_Summary[i]
        #replace each <p> <ul> to </p> <ul> <br>
        temp_string <- all_scenarios_df3$Installation_Summary[i]
        temp_string1 <- stringr::str_replace_all(temp_string, "</p> <ul>", "</p> <ul> <br>")
        all_scenarios_df3$Installation_Summary[i] <- temp_string1
      }
  
# add full SITENAME, SITEID ---- *********************************************MA CHANGED THIS TO INCLUDE SITENAME
  siteid_string <- all_scenarios_df3$SITEID[1]
  siteid_string1 <- stringr::str_replace_all(siteid_string, "<p>", '')
  siteid_string1 <- stringr::str_replace_all(siteid_string1, "</p>", '')
  all_scenarios_df3$SITEID <- siteid_string1 
  
  SITENAME <- installation_info[installation_info$SITEID == all_scenarios_df3$SITEID[1], 2]
  
  if(!is.na(SITENAME)){
    all_scenarios_df3[,"SITENAME"] <- SITENAME
  }else(print("No match found in installation database"))


  
# ADDING NAVY VIEWER CODE #########################
  
  if(inst_sheet == "NAVY"){
    sample_doc <- all_scenarios_df3
    
    sample_doc1 <- sample_doc %>% 
      select(-c(6:13)) %>%
      mutate(scenario_period = paste0(Scenario, "/", Period),
             scenario_period = recode(scenario_period, "Historical/Historical" = "Modeled Historical Baseline")) %>% 
      relocate(scenario_period, .before = Scenario) %>% 
      select(-Scenario, -Period, -SITEID)
    
    sample_doc_long <- sample_doc1 %>% 
      pivot_longer(cols = names(sample_doc1)[2:9])
    
    # Trying to attach scenario to each
    
    new_names <- c(
      "Modeled Historical Baseline",
      "Moderate Disruption/Near Term",
      "Moderate Disruption/Far Term",
      "High Disruption/Near Term",
      "High Disruption/Far Term"
    )
    
    sample_doc_long2 <- sample_doc_long %>%
      mutate(
        Scenario = ifelse(value %in% new_names, value, NA)
      ) %>%
      fill(Scenario, .direction = "down")
    
    
    ######### Set up "Chart"
    
    chart <- c("SPEI_Text", #1
               "SPEI", #2
               "Dry_Distribution_Text", #3
               "Dist", #4
               "Wet_Distribution_Text", #5
               "Dry_Duration_Severity_Text", #6
               "Dur Sev", #7
               "Wet_Duration_Severity_Text") #8
    
    sample_doc_long3 <- sample_doc_long2 %>%
      mutate(
        Chart = ifelse(name %in% chart[1], chart[2], NA),
        Chart = ifelse(name %in% chart[3], chart[4], Chart),
        Chart = ifelse(name %in% chart[5], chart[4], Chart),
        Chart = ifelse(name %in% chart[6], chart[7], Chart),
        Chart = ifelse(name %in% chart[8], chart[7], Chart),
        Chart = ifelse(is.na(Chart), "Summary", Chart)
      )
    
    
    ######## Set up "Condition"
    
    condition <- c("Dry", "Wet")
    
    sample_doc_long4 <- sample_doc_long3 %>% 
      mutate(Condition = ifelse(name %in% chart[3] | name %in% chart[6], condition[1],
                                ifelse(name %in% chart[5] | name %in% chart[8], condition[2], 
                                       ifelse(Scenario == "Modeled Historical Baseline", "Modeled Historical Baseline",
                                              ifelse(Scenario == "Moderate Disruption/Near Term", "Moderate Disruption/Near Term",
                                                     ifelse(Scenario == "Moderate Disruption/Far Term", "Moderate Disruption/Far Term",
                                                            ifelse(Scenario == "High Disruption/Near Term", "High Disruption/Near Term",
                                                                   ifelse(Scenario == "High Disruption/Far Term", "High Disruption/Far Term", NA))))))),
             Condition = ifelse(name == "Installation_Summary", "Installation Summary",
                                ifelse(name == "References", "References", Condition)))
    
    ### Adding the correct ordering
    Scenario_levels <- c("Moderate Disruption/Near Term", "Moderate Disruption/Far Term", "High Disruption/Near Term", "High Disruption/Far Term", "Modeled Historical Baseline")
    Chart_levels <- c("SPEI", "Dur Sev", "Dist", "Summary")
    # Condition_levels <- c("Moderate Disruption/Near Term", "Moderate Disruption/Far Term", "High Disruption/Near Term", "High Disruption/Far Term", "Dry", "Wet", "")
    
    sample_doc_long5 <- sample_doc_long4 %>% 
      mutate(Scenario = factor(Scenario, levels = Scenario_levels),
             Chart = factor(Chart, levels = Chart_levels)) %>% 
      arrange(Chart, Scenario, Condition) %>% 
      relocate(c(Scenario, Chart, Condition), .after = SITENAME) %>% 
      select(-name)
    
    ##### Get rid of NA text fields and unnecessary Historical Baseline repeats
    sample_doc_long6 <- sample_doc_long5 %>% 
      filter(!(Scenario == "Modeled Historical Baseline" & Chart == "Dur Sev"),
             !(Scenario == "Modeled Historical Baseline" & Chart == "Dist"),
             !(Chart == "Summary" & is.na(value)),
             !(Chart == "Summary" & Condition == "Modeled Historical Baseline"),
             !(Chart == "Summary" & str_detect(Scenario, "Disruption")))
    
    ##### Change titles of some columns
    sample_doc_long6 <- sample_doc_long6 %>% 
      rename(Installation = SITENAME,
             Text = value)
    
    
    ##### Add a bunch of historical columns
    row_to_copy <- sample_doc_long6 %>% slice(5)
    
    sample_doc_long7 <- sample_doc_long6 %>% 
      mutate(row_id = row_number(),
             Scenario = if_else(row_number() == 5, sample_doc_long6$Scenario[4], Scenario)) %>% 
      bind_rows(row_to_copy %>% mutate(row_id = 1.5, Scenario = sample_doc_long6$Scenario[1]),
                row_to_copy %>% mutate(row_id = 2.5, Scenario = sample_doc_long6$Scenario[2]),
                row_to_copy %>% mutate(row_id = 3.5, Scenario = sample_doc_long6$Scenario[3])) %>% 
      arrange(row_id) %>% 
      select(-row_id)
    
    ##### Try to separate the Summary column into 2
    
    row <- sample_doc_long7[25, ]
    
    part1 <- row
    part1$header <- "Precipitation and Drought"
    part1$Text <- str_trim(
      str_replace(
        part1$Text,
        "<p>Precipitation and Drought</p>\\s*<ul>\\s*<br>",
        "Precipitation and Drought<ul>"
      )
    )
    part1$Text <- str_trim(
      str_remove(
        str_extract(
          part1$Text,
          regex(
            "Precipitation and Drought.*?(?=Other Water-Related Considerations)",
            dotall = TRUE
          )
        ),
        "^Precipitation and Drought\\s*"
      )
    )
    
    
    part2 <- row
    part2$header <- "Other Considerations"
    part2$Text <- str_replace(
      part2$Text,
      "<p>Other Water-Related Considerations</p>\\s*<ul>\\s*<br>",
      "Other Water-Related Considerations<ul>"
    )
    part2$Text <- str_trim(
      str_remove(
        str_extract(
          part2$Text,
          regex(
            "Other Water-Related Considerations.*",
            dotall = TRUE
          )
        ),
        "^Other Water-Related Considerations\\s*"
      )
    )
    
    
    df <- bind_rows(
      sample_doc_long7[-25, ],
      part1,
      part2
    )
    
    ##### Now I'm looking at the final arrangement of data 
    
    df2 <- df %>% 
      mutate(Condition = case_when(
        !is.na(header) ~ header, TRUE ~ Condition)) %>% 
      select(-header)
    
    
    summary_rows <- df2 %>%
      filter(Chart == "Summary") %>%
      mutate(condition_order = match(Condition, c("Precipitation and Drought",
                                                  "Other Considerations",
                                                  "References")),
             Scenario = "") %>% 
      arrange(condition_order) %>%
      select(-condition_order)
    
    df3 <- bind_rows(
      filter(df2, Chart != "Summary"),
      summary_rows
    )
  }
  
  
  # Export final files ----
  ##export excel to 3ViewerPackages folder ----
  
  # Figure out what output to use (AF vs NAVY) ############# MA ADDED THIS TO MAKE SURE THE RIGHT OUTPUT COMES OUT, ALSO FILENAME HAS CHANGED TO REFLECT THIS
  ifelse(inst_sheet == "NAVY", output_file <- df3, output_file <- all_scenarios_df3)
  
  
  out_dir <- paste0(input_umbrella, input_installation_folder, "/3ViewerPackages/HTML_excels") 
  # ******** NOTE THAT THE FOLDER STRUCTURE MUST MATCH WHAT IS ABOVE ^^^ EXACTLY.  **********
  # CHANGE out_dir AS NEEDED IF THERE ARE ANY DIFFERENCES IN THE LOCATION YOU WANT TO SAVE TO.
  
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  
  output_filename <- paste0(project_name, "_HTML_formatted_", current_date, ".xlsx")
  write_xlsx(output_file, file.path(out_dir, output_filename)) #create file and save to 3ViewerPackages folder
  message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))
  
  ##create shortcut to Word to HTML folder ----
  out_full_path <- file.path(out_dir, output_filename) #save the path to the excel in 3ViewerPackages
  output_filelink <- paste0(project_name, "_HTML_formatted_", current_date, ".lnk") #create shortcut name
  shortcut_location <- file.path(input_dir, output_filelink) #save the path to the future shortcut
  
  shell(paste0( #create shortcut to Word to HTML Conversion folder (this uses the Windows power shell)
    'powershell -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; ',
    '$s = $ws.CreateShortcut(\'', shortcut_location, '\'); ',
    '$s.TargetPath = \'', out_full_path, '\'; ',
    '$s.Save()"'
  )) 

  
  