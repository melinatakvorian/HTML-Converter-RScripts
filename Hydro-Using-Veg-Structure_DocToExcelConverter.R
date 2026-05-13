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
  #AIR FORCE  
  #input_umbrella <- "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/" 
  
  #NAVY
  input_umbrella <- "N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/"
  
  #the specific folder inside the Document to HTML Table Converter where the input files are
  input_specific_folder <- "MidLant Region/WPNSTA Yorktown/Hydrology/HTML converter" 
  
  #the final file name will start with this and will get the date added
  subject <- "Hydro"
  installation <- "Yorktown"
  project_name <- paste0(subject, "_", installation) #Replace with whatever you want.

#####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

input_dir <-  paste0(input_umbrella, input_specific_folder) #Rename to your target directory. Outputs will appear here as well.
current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"
installation_info <- readxl::read_xlsx("N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/_AirForceClimateViewerDev/Document to HTML Table Converter/FilesForTesting/Installation_Info.xlsx")

#ERROR CATCH 1 ----

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


# ----- *Word->HTML function ----
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


# ----- *HTML->pieces function ----
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
# ----- *removing spaces after headings function -----
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

# ----- *replace the last instance of a substring -----
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
  
  # Define indices for histclimatic and disruptions sections
  hist_indices <- c(1:6, last) # histclimatic sections
  disr_indices <- c(1:2, 7:(last-1)) # disruptions sections
  
  
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


#unfold the results list to be able to create a dataframe
all_headings_hist <- unique(unlist(lapply(results_hist, names)))
all_headings_disr <- c("SITENAME", "SITEID")
all_headings_disr <- append(all_headings_disr, unique(unlist(lapply(results_disr, names))))

# Create dataframe and input HTML in proper sections ----

##hist----
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

##DISRUPTION SCENARIOS----

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

for(file in seq_along(results_disr)){
  for(a in seq_along(mylist[[file]])){
    # Extract the current list of indices from mylist
    templist <- mylist[[file]][[a]]
    
    # Populate the first few columns with results_hist data (assuming it applies to all rows for this file)
    df_disr[rownum, 1] <- results_hist[[file]][[1]]
    df_disr[rownum, 2] <- results_hist[[file]][[2]]
    
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

##Delete empty columns ----
test <- df_disr
# empty_cols <- c()
# 
# for(i in 1:ncol(test)){
#   if(all(is.na(test[[i]]))){
#     empty_cols[length(empty_cols)+1] <- i
#   }
# }
# 
# df_disr <- df_disr[ , -empty_cols]

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
    frame <- matrix(nrow = 5, ncol = length(colnames))
    frame <- as.data.frame(frame)
    colnames(frame) <- colnames
    frame$Scenario <- scenario
    frame$Period <- period
  
    frame2 <- frame
  
  ###start transposing data starting AFTER historical row----
    inst_summ <- as.numeric(which(colnames(df_hist) == "Installation_Summary"))
    
    #Fill SITENAME, SITEID, Installation Summary
    frame2$Installation_Summary <- df_hist[1,inst_summ]
    frame2$SITENAME <- df_hist$SITENAME
    frame2$SITEID <- df_hist$SITEID
    
    #Fill row 1 with historical data
    frame2$SPEI_Text[1] <- df_hist$`Period: Historical, SPEI_Text`[1]
    frame2$References[1] <- df_hist$References[1]
    
    
    # Fill rows 2,4 for NEAR TERM scenarios
    frame2$SPEI_Text[[2]] <- df_near_term$`Period: Near Term, SPEI_Text`[1]
    frame2$SPEI_Text[[4]] <- df_near_term$`Period: Near Term, SPEI_Text`[2]

    
    # Fill rows 3,5 for FAR TERM scenarios
    frame2$SPEI_Text[[3]] <- df_far_term$`Period: Far Term, SPEI_Text`[1]
    frame2$SPEI_Text[[5]] <- df_far_term$`Period: Far Term, SPEI_Text`[2]

    
    #Fill in the other rows based on disruption
    frame2[c(4,5), 7:10] <- leftovers[2, 5:8] #high disruption
    frame2[c(2,3), 7:10] <- leftovers[1, 5:8] #moderate disruption
    
#add BLANK numeric columns----
  
  cols <- as.numeric(ncol(frame2))+1
  cols_w_nos <- cols + 7
  frame2[,c(cols:cols_w_nos)] <- ""
  
  #numeric column names
  new_cols <- c("Minimum_SPEI", "Maximum_SPEI", 
                "Dry_Variability", "Wet_Variability", "Dry_Events", "Wet_Events", 
                "Dry_Change", "Wet_Change", "")
  
  #assign names
  colnames(frame2)[cols:cols_w_nos] <- new_cols #this one errors, don't worry about it
  
  #move columns to where Anthony wants them
  frame2 <- frame2 %>% 
    relocate(all_of(cols:cols_w_nos), .before = "Installation_Summary")
    
  frame3 <- frame2
  
#ADDING INDENTS AND LINE BREAKS----
  #add REFERENCES SECTION HANGING INDENT
  for(i in 1:nrow(frame2)){
    if(is.na(frame2$References[i])) next #skip NA rows
    
    frame2$References[i]
    
    #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
    temp_string <- frame2$References[i]
    
    #temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style="padding-left:15px;text-indent:-15px;">')
    temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
    temp_string2 <- replace_all_except_last(temp_string1, "</p>", "</p> <br>")
    frame2$References[i] <- temp_string2
  }
  
  #these are the sections that need the edits
  #"SPEI_Text", "Dry_Distribution_Text", "Wet_Distribution_Text", "Dry_Duration_Severity_Text", "Wet_Duration_Severity_Text"
    
  numbblocks <- c(5, 15:18) #corresponds to the columns with text that we need broken up
  
    #add blank line after each paragraph
    for(a in 1:length(numbblocks)){
      col_num <- numbblocks[[a]]
      for(b in 1:nrow(frame2)){
        if(is.na(frame2[[col_num]][b])) next
        
        #replace each </p> to </p> <br>
        temp_string <- frame2[[col_num]][b]
        temp_string1 <- replace_all_except_last(temp_string, "</p>", "</p> <br>")
        frame2[[col_num]][b] <- temp_string1
      }
    }
  
    #add indent at the beginning of each non-bulleted paragraph
    for(a in 1:length(numbblocks)){
      col_num <- numbblocks[[a]]
      for(b in 1:nrow(frame2)){
        
        if(is.na(frame2[[col_num]][b])) next
        
        #replace each <p> to <p style=text-indent:-15px;>
        temp_string <- frame2[[col_num]][b]
        temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=text-indent:15px;>')
        frame2[[col_num]][b] <- temp_string1
      }
    }
  
    #add blank line after each bulleted paragraph
    for(i in 1:nrow(frame2)){
      if(is.na(frame2$Installation_Summary[i])) next
      
      #replace each </p></li> to </p></li><br>
      temp_string <- frame2$Installation_Summary[i]
      temp_string1 <- replace_all_except_last(temp_string, "</p></li>", "</p></li><br>")
      frame2$Installation_Summary[i] <- temp_string1
    }
  
    #add blank line after the subheading in Installation Summary
    for(i in 1:nrow(frame2)){
      frame2$Installation_Summary[i]
      #replace each <p> <ul> to </p> <ul> <br>
      temp_string <- frame2$Installation_Summary[i]
      temp_string1 <- stringr::str_replace_all(temp_string, "</p> <ul>", "</p> <ul> <br>")
      frame2$Installation_Summary[i] <- temp_string1
    }
  
  # add full SITENAME, SITEID ----
  key <- match(installation, installation_info$ShortName)
  
  if(!is.na(key)){
    frame2[,"SITENAME"] <- installation_info$SITENAME[key]
    frame2[,"SITEID"] <- installation_info$SITEID[key]
  }else(print("No match found in installation database"))
    
# Export final files ----
out_dir <- input_dir
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

#final file
output_filename <- paste0(project_name, "_HTML_formatted_", current_date, ".xlsx")
write_xlsx(frame2, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment so that things can run properly for the next run  
#rm(list = ls()) 

