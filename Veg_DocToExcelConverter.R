# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2025.11.14


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

  packages <- c("pandoc","xml2","rvest","writexl", "readxl", "tidyverse")
  
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
  input_installation_folder <- "WPNSTA Yorktown"
  # Write if working on AF (AIR FORCE) or Navy (NAVY):
  inst_sheet = "NAVY"
  # inst_sheet = "AIR FORCE"
  # If Navy, select which region
  navy_region = "MidLant Region"
  # navy_region = "Southeast Region"
  # navy_region = "Hawaii Region"
  # Select which analysis you're doing (shouldn't need to change)
  input_SME_folder <- "/Vegetation_Habitats/Word to HTML Conversion/TEST" 
  #the final file name will start with this and will get the date added
  subject <- "Veg"
  project_name <- paste0(subject, "_", input_installation_folder)
  # this will select which sheet to select your data from
  ifelse(inst_sheet == "AIR FORCE",
         input_umbrella <- "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/",
         input_umbrella <- paste0("N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/", navy_region, "/"))
  
  
    
#####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

  input_dir <-  paste0(input_umbrella, input_installation_folder, input_SME_folder) 
  current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"
  installation_info <- readxl::read_xlsx("Installation_IDs.xlsx", sheet = inst_sheet)
  

#ERROR CATCH: open files ----

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
      standalone = TRUE,
      args = c("--wrap=none") #change to preserve?
    )
    
    xml2::read_html(html_file)
  }


# ----- * HTML->pieces function ----
#reads HTML file (input) and separate sections for building table later
  parse_html_sections_bio <- function(html_doc, section_indices) {
    #identify all headings
    headings <- rvest::html_nodes(html_doc, "h1") #identify headings
    sections <- vector("list", length(section_indices)) #create list of headings (sections)
    
    for (i in seq_along(section_indices)) { # Iterate over specified sections
        print(paste("Parsing section:", section_indices[i]))
        
      start_node <- headings[[section_indices[i]]]
      
      # MA - ADDED THIS 9/23 TO FIX ISSUES WITH HEADERS ACCIDENTALLY BEING INCLUDED IN OTHER SECTIONS
      end_node <- if (section_indices[i] < length(headings))
        headings[[section_indices[i] + 1]]
      else
        NULL
      
      # end_node <- if (i < length(section_indices)) headings[[section_indices[i + 1]]] else NULL
      # end_node <- if (i < length(section_indices)) headings[[i + 1]] else NULL
        #print(headings[section_indices[i + 1]])
      print(section_indices)

      # print(xml_text(headings[[section_indices[i]]]))
      # print(xml_text(headings[[section_indices[i + 1]]]))

      siblings <- xml2::xml_find_all(start_node, "following-sibling::*")
      if (!is.null(end_node)) {
        idx <- which(vapply(siblings, identical, logical(1), y = end_node))
        if (length(idx) == 0) idx <- length(siblings) + 1
        siblings <- siblings[seq_len(idx - 1)]
      }
      
      # Insert a space between concatenated HTML nodes
      content_html <- paste(as.character(siblings), collapse = " ")
      sections[[i]] <- content_html
        #print(content_html)
    }
    
    # Assign section titles as names to the list elements
    names(sections) <- sapply(headings[section_indices], xml_text)
    sections
    
  }
  
  
  parse_html_sections_veg <- function(html_doc, section_indices) {
    #identify all headings
    headings <- rvest::html_nodes(html_doc, "h1") #identify headings
    sections <- vector("list", length(section_indices)) #create list of headings (sections)
    
    for (i in seq_along(section_indices)) { # Iterate over specified sections
      print(paste("Parsing section:", section_indices[i]))
      
      start_node <- headings[[section_indices[i]]]
      
      # MA - ADDED THIS 9/23 TO FIX ISSUES WITH HEADERS ACCIDENTALLY BEING INCLUDED IN OTHER SECTIONS
      end_node <- if (section_indices[i] < length(headings))
        headings[[section_indices[i] + 1]]
      else
        NULL
      # end_node <- if (i < length(section_indices)) headings[[section_indices[i + 1]]] else NULL
      #end_node <- if (i < length(section_indices)) headings[[i + 1]] else NULL
        #print(headings[section_indices[i + 1]])
      
      siblings <- xml2::xml_find_all(start_node, "following-sibling::*") #finds all instances of the heading in the document
      if (!is.null(end_node)) { 
        idx <- which(vapply(siblings, identical, logical(1), y = end_node)) #
        if (length(idx) == 0) idx <- length(siblings) + 1
        siblings <- siblings[seq_len(idx - 1)]
      }
      
      # Insert a space between concatenated HTML nodes
      content_html <- paste(as.character(siblings), collapse = " ")
      sections[[i]] <- content_html
      #print(content_html)
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

  # ----- * remove paragraph notation -----
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
  
  
# ----- * remove '\r\n' from heading names -----
  
  #NO LONGER NECESSARY, since we have removed the text wrapping default from the pandoc_convert() function
  
  #if results[i] includes '\r\n', remove it
  # remove_accidental_return <- function(result_list){
  #   
  #   for(i in 1:length(result_list)){
  #     templist <- result_list[[i]]
  #     
  #     for(heading in 1:length(templist)){
  #       if(stringr::str_detect(names(templist)[heading], "\\r\\n")){
  #         
  #         #replace "\r\n" with nothing
  #         headingWithProblem <- names(templist)[heading] #save heading to local object
  #         
  #         newHeading <- stringr::str_replace_all(headingWithProblem, "\\r\\n", " ")
  #         
  #         names(result_list[[i]])[heading] <- newHeading
  #         print(names(result_list[[i]][heading]))
  #       }else next
  #     }
  #   }
  #   return(result_list)
  # }


# RUN ----
  
  #initialize objects for storing file info
  docx_files <- list.files(input_dir, pattern = "\\.docx$", full.names = TRUE) #pull list of all files in folder
  results_bio <- list()
  results_veg <- list()
  
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
      
      # Define indices for bioclimatic and vegetation sections
      # separated by Navy and Air Force

      if(inst_sheet == "NAVY"){
        bio_indices <- c(1:4, (last-1):last) #this is installation
        veg_indices <- c(1, 5:(last-2)) # this is group
      } else {
        bio_indices <- c(1:8, last) # Bioclimatic sections
        veg_indices <- c(1:3, 9:(last - 1)) # Vegetation sections
      }
      
      #Create BIO table list
      sections_bio <- parse_html_sections_bio(html_doc, bio_indices)
      sections_bio <- sections_bio[names(sections_bio) != ""] #remove accidental headers
      results_bio[[basename(file)]] <- sections_bio #should be a list of headings and its text
      
      #Create VEG table list
      sections_veg <- parse_html_sections_veg(html_doc, veg_indices)
      sections_veg <- sections_veg[names(sections_veg) != ""] #remove accidental headers
      results_veg[[basename(file)]] <- sections_veg #should be a list of headings and its text
    }
  
  
#remove blank spaces after headings that could cause additional headers accidentally
  results_bio <- remove_end_blanks(results_bio)
  results_veg <- remove_end_blanks(results_veg)
  #results_veg <- remove_accidental_return(results_veg) #see the function above for why this is commented out


#unfold the results list to be able to create a dataframe
  all_headings_bio <- unique(unlist(lapply(results_bio, names)))
  all_headings_veg <- unique(unlist(lapply(results_veg, names)))
  
# Create dataframe and input HTML in proper sections ----
  
  ##BIO----
    df_bio <- data.frame(matrix(NA_character_, length(results_bio), length(all_headings_bio)),
                         stringsAsFactors = FALSE)
    colnames(df_bio) <- all_headings_bio
    rownames(df_bio) <- names(results_bio)
    
    for (i in seq_along(results_bio)) {
      for (col in all_headings_bio) {
        if (col %in% names(results_bio[[i]])) {
          df_bio[i, col] <- results_bio[[i]][[col]]
        }else{df_bio[i, col] <- NA} #ChatGPT help
      }
    }
    
  ##VEG----

    #find the indices within the list that are new occurrences of 'Vegetation Group Name'
      num_files <- as.list(c(1:as.numeric(length(results_veg)))) #initialize list
      
      #create mini lists for each instance of new veg group
        for(file in seq_along(results_veg)){
          veg_names <- names(results_veg[[file]]) #create list of headings from each file
          indices <- c()
          
          for(i in seq_along(veg_names)){ 
            if(grepl("G[0-9]+", veg_names[i])){  #if the heading at this index is a new veg group
              indices[length(indices)+1] <- i} #save the index to the end of the indices list
          }
          num_files[[file]] <- indices #append the indices of new veg group to this number file in the folder
        }

    #use the indices to create smaller lists as keys to sections of Veg Groups in the document
      #initialize objects
      total_rows <- 0
      split_sections <- vector("list", length(results_veg))
  
      #create mini lists, assign data to them
      for(file in seq_along(results_veg)){ #for each heading
        
        for(i in seq_along(num_files[[file]])){ #in each file
          finish <- i+1 
          
          if(finish <= length(num_files[[file]])){ #if we aren't past the last file in the folder,
            secondtolast <- num_files[[file]][[finish]] #get the instance of the next New Veg heading
            secondtolast <- secondtolast-1 #we want to stop BEFORE we get to the next section
            
            num_pair <- c(num_files[[file]][[i]]:secondtolast) #create range from one to the next
            
            total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
            
            split_sections[[file]][[length(split_sections[[file]])+1]] <- num_pair #create a nested list with each index within a veg group section
            
          }else{ #case for the last instance of new veg group

            num_pair <- c(num_files[[file]][[i]]:(length(results_veg[[file]]))) #trying without the subtract 1
              #MT- this used to say "num_pair <- c(num_files[[file]][[i]]:(length(results_veg[[file]]-1)))".
                #...by removing the '-1', this includes the last index of the last heading of this section
                #...which resolves an issue where split_sections did not include all the last heading's index
            
            total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
            
            split_sections[[file]][[length(split_sections[[file]])+1]] <- num_pair
          }
        }
      }

    #create a df where each row is one of these lists. 
      # ------------------------------------------------------------------
      #MT - right now, the way the number of rows is determined for df_veg is nrow=length(total_rows), but total_rows = 32, instead of 3 for this dummy dataset.
      #MT - df_veg is also only one row, despite trying to be created with nrow = 32. I will change this to nrow = length(num_files[[1]], 
      #...which corresponds to the number of veg groups identified in the process in the code chunk above)
      #MT - this adjustment worked for nrow!
      
      #MT- I need to also adjust the number of columns and what is going into the column names. Right now ncol=length(unique(all_headings_veg))
      #...but this includes the names of the new veg groups (G####, G###, etc.). We need a new object to store the headings, that excludes the G### strings
      
      unique_headings_veg <- c()
      veg_group_names <- c()
      for(i in seq_along(all_headings_veg)){ 
        if(grepl("G[0-9]+", all_headings_veg[i])){ #if the heading at this index IS a new veg group
          veg_group_names[length(veg_group_names)+1] <- all_headings_veg[i]} #save this string to the veg_group_names object
        
        if(!grepl("G[0-9]+", all_headings_veg[i])){  #if the heading at this index is NOT a new veg group
          unique_headings_veg[length(unique_headings_veg)+1] <- all_headings_veg[i]} #save this string to the end of the unique_headings_veg object
      }
      
      #MT - For now, I think we should not worry about getting the veg group names into a column. 
        #...I think it makes sense to make sure the data will be inserted correctly, then worry about that.
      
      df_veg <- data.frame(matrix(NA_character_, nrow=length(num_files[[1]]), ncol=length(unique_headings_veg)),
                           stringsAsFactors = FALSE)

      
      colnames(df_veg) <- unique_headings_veg
      rownum <- 1
      
      #MT- I am noticing an issue where the data is being input into SITEID correctly, but the rest is one column to the right.
        #...if n_col is set to 1, then SITEID is not populated. if n_col is set to 2, then SITEID is populated, but the rest of the data is offset by one.  
        #...This must have to do with the order that the new SITEID column is being incorporated, so I am going to move things around to see what works.
        #...moving the SITEID to go after the veg group data worked! Now SITEID AND veg group data are in the df, as desired.
      
      #MT- The last issue is that there is the last cell of the df is not being populated. It is because split_sections cuts off the last header of the last veg group. 
        #...This needs to be worked on where split_sections is created.
        #...I found out how to get the split_sections to include the last heading in the veg group! see the edits made up there.

      for(file in seq_along(results_veg)){
        for(a in seq_along(split_sections[[file]])){
          # Extract the current list of indices from split_sections
          templist <- split_sections[[file]][[a]]
          
          n_col <- 1 # Start filling from the 2nd column
          
          # Extract elements from results_veg based on the indices in templist
          for(b in seq_along(templist)){
            df_veg[rownum, n_col] <- results_veg[[file]][[templist[[b]]]]
            n_col <- n_col + 1
          }
          
          # Populate the first column with results_bio data (assuming it applies to all rows for this file)
          df_veg[rownum, 1] <- results_bio[[file]][[1]] #fills in the SITEID information
          
          # Move to the next row for the dataframe
          rownum <- rownum + 1
        }
      }

  ##Create VegGroup column and add name ----
      
      for(row in 1:nrow(df_veg)){
        df_veg$VegGroup[row] <- veg_group_names[[row]]
      }
      
      df_veg <- df_veg %>% relocate(VegGroup, .before = 'Group_Desc')
        
  ##Extract group name and groupNum from VegGroup
      df_veg <- df_veg %>%         # This will split VegGroup into 2, removing the VegGroup column
        separate_wider_regex(
          cols = VegGroup,
          patterns = c(             # this is code from Copilot!
            GroupNum = "^G\\d+",    # Captures G and the numbers and the colon
            ": ",                   # Drops the semicolon and space in between
            GroupName = ".*"        # Captures everything else
          ),
          too_few = "align_start"   # Prevents errors if a row doesn't match perfectly
        )

  ##Delete empty columns ----
      #THIS IS PROBABLY NOT NECESSARY ANYMORE. It was originally made to handle columns that were just there as tags, without data
    # test <- df_veg
    # empty_cols <- c()
    #   
    # for(i in 1:ncol(test)){
    #   if(all(is.na(test[[i]]))){
    #     empty_cols[length(empty_cols)+1] <- i
    #   }else if(all(test[[i]] == "")){
    #     empty_cols[length(empty_cols)+1] <- i
    #   }
    # }
    #   
    # df_veg <- df_veg[ , -empty_cols]

## MA - The below section won't be needed for Navy
#make Exposure Icon column for Anthony
# df_veg[, 'Exposure_Icon'] <- "Extreme Heat, Drought, Vector Borne Disease, Invasive Species, Seasonal Timing, Fire/Flooding"
# 
      
  ##references hanging indent ----
    #add REFERENCES SECTION HANGING INDENT <p style=padding-left:15px;text-indent:-15px;>
    for(i in 1:nrow(df_bio)){
      df_bio$References[i]
      #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
      temp_string1 <- df_bio$References[i]
      temp_string2 <- stringr::str_replace_all(temp_string1, "<p>", "<p style=padding-left:15px;text-indent:-15px;>")
      df_bio$References[i] <- temp_string2
    }

# add full SITENAME, SITEID ----
  
  # MA - removing paragraph notation for both veg and bio dataframes
  cols_to_change <- c("SITEID") #change this to the name of the columns in the specific analysis
  df_bio <- p_be_gone(df_bio, cols_to_change)
  df_veg <- p_be_gone(df_veg, cols_to_change)
  
  # Adding SITENAME and InstallationID
  SITENAME <- installation_info$InstallationNames[installation_info$SITEID == df_bio$SITEID[1]]
  InstallationID <- installation_info$InstallationID[installation_info$SITEID == df_bio$SITEID[1]]
  
  df_bio$InstallationName <- SITENAME
  df_bio$InstallationID <- InstallationID
  df_veg$InstallationName <- SITENAME
  df_veg$InstallationID <- InstallationID
  
  df_veg <- df_veg %>% relocate(c(InstallationID, InstallationName), .before = `GroupNum`)
  df_bio <- df_bio %>% relocate(c(InstallationID, InstallationName), .after = `SITEID`)

  ### QUESTION FOR KT OR MT - how can the above step be made better so it doesn't take like 12 lines of code?
  
# Creating Selector csv ----
  
  ##SELECTOR FOR LIST ITEMS
  # Find the Exposure Description header in the doc
  desc_header <- html_nodes(html_doc, "h1") %>%
    .[html_text(.) == "ExpDescription"]
  
  # Get the list after the header
  desc_list <- xml2::xml_find_all(desc_header, "following-sibling::ol")
  
  # Extract list items
  bioclimatic5 <- html_nodes(desc_list, "li")
  bio5_text <- html_text(bioclimatic5, trim = TRUE)
  
  # Get Julia's table from the selector_topfive worksheet
  lookup_Top5 <- read_excel(
    "N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/_SMEs Dashboard Dev/Dashboard_inputs/5_TerrVeg/Notes for Terrestrial Vegetation Dashboard inputs.xlsx",
    sheet = "selector_topfive",
    range = "D9:F26",
    col_names = c("Top5_Code", "Top5_Variable_GIS", "LongFormName_Veg")
  )
  
  # Match the table items to the list items
  results_Top5 <- data.frame(
    LongFormName_Veg = bio5_text,
    stringsAsFactors = FALSE
  ) %>%
    left_join(lookup_Top5 %>% select(LongFormName_Veg), by = "LongFormName_Veg") %>% 
    mutate(SiteID = df_bio$SITEID,
           InstallationID = df_bio$InstallationID,
           InstallationName = df_bio$InstallationName)

  # assign code values to the bioclimatic variables
    results_Top5 <- results_Top5 %>% mutate('Top5_Code' = case_when(
      LongFormName_Veg == "Annual Mean Diurnal Range, °F" ~ 1,
      LongFormName_Veg == "Isothermality, %"  ~ 2,
      LongFormName_Veg == "Temperature Seasonality (Standard Deviation), °F" ~ 3,
      LongFormName_Veg == "Temperature Seasonality (Coefficient of Variation), %" ~ 4,
      LongFormName_Veg == "Max Temperature of Warmest Month, °F" ~ 5,
      LongFormName_Veg == "Min Temperature of Coldest Month, °F" ~ 6,
      LongFormName_Veg == "Annual Temperature Range, °F" ~ 7,
      LongFormName_Veg == "Mean Temperature of Wettest Quarter, °F" ~ 8,
      LongFormName_Veg == "Mean Temperature of Driest Quarter, °F"~ 9,
      LongFormName_Veg == "Mean Temperature of Warmest Quarter, °F" ~ 10,
      LongFormName_Veg == "Mean Temperature of Coldest Quarter, °F" ~ 11,
      LongFormName_Veg == "Precipitation of Wettest Month, inches" ~ 12,
      LongFormName_Veg == "Precipitation of Driest Month, inches" ~ 13,
      LongFormName_Veg == "Precipitation Seasonality (Coefficient of Variation), %" ~ 14,
      LongFormName_Veg == "Precipitation of Wettest Quarter, inches" ~ 15,
      LongFormName_Veg == "Precipitation of Driest Quarter, inches" ~ 16,
      LongFormName_Veg == "Precipitation of Coldest Quarter, inches" ~ 17,
      LongFormName_Veg == "Precipitation of Warmest Quarter, inches" ~ 18,
      TRUE ~ 0,
    ))
    
    results_Top5 <- results_Top5 %>% 
      left_join(lookup_Top5 %>% 
                  select(Top5_Code, Top5_Variable_GIS), by = "Top5_Code")
    
    results_Top5 <- results_Top5[c(2,3,4,1,5,6)] # Reorganizing the columns for Julia's template
    
# Creating Group Description csv ----
  last_grp <- length(df_veg)
  grp_dsc_indices <- c(1:6, (last_grp-1):last_grp)
  
  group_desc <- df_veg %>% 
    select(grp_dsc_indices)
  

# Creating Group Icons csv ----

  #delineating category:
  sensitivity <- c("Landscape Condition", "Fire", "Insects and Disease", "Invasive and Ruderal Vegetation")
  adaptive_capacity <- c("Topoclimatic Variability", "Diversity within Functional Species Groups", "Keystone Species Vulnerability")

  # As a test, using what I had in my test dataframe
  # sensitivity <- c("Landscape condition", "Fire", "Insects and Disease", "Invasive and ruderal vegetation")
  # adaptive_capacity <- c("Topoclimatic variability", "Diversity within functional species groups", "Keystones species vulnerability")
  
  # preparing the dataframe for pivot longer
  df_group_icons <- df_veg %>% 
    select(-c(6, 14, 15)) # these are hard coded based on the template word doc
  
  # pivoting longer
  df_group_icons_l <- pivot_longer(
    df_group_icons,
    cols = -c(1:5),
    names_to = "Icon Name",
    values_to = "Text"
  )
  
  # adding category, numeric value
  df_group_icons_l <- df_group_icons_l %>% 
    mutate(
      Category = ifelse(`Icon Name` %in% sensitivity, "Sensitivity", "Adaptive Capacity"),
      Code = case_when(`Icon Name` == sensitivity[1] ~ 1,
                       `Icon Name` == sensitivity[2] ~ 2,
                       `Icon Name` == sensitivity[3] ~ 3,
                       `Icon Name` == sensitivity[4] ~ 4,
                       `Icon Name` == adaptive_capacity[1] ~ 5,
                       `Icon Name` == adaptive_capacity[2] ~ 6,
                       `Icon Name` == adaptive_capacity[3] ~ 7)
    ) %>% 
    relocate(Code, .before = `Icon Name`) %>% 
    relocate(Category, .before = Text)

# Creating Installation csv ----
  rownames(df_bio) <- 1
  df_installation <- df_bio %>% relocate(NotAnalyzed, .after = 'References')
  ### NOTE - We probably need to go into this df and make sure adequate "breaks" are included!
  
# Export final files ----
  
  # MA - note on final export files for Navy:
    # Selector CSV - results_Top5
    # Group Description CSV - group_desc
    # Group Icons csv - df_group_icons_l
    # Installation csv - df_installation
  
  ### MA note - I haven't played with exporting!
  
  ##export excel to 3ViewerPackages folder ----
    out_dir <- paste0(input_umbrella, input_installation_folder, input_SME_folder, "/3ViewerPackages/TEST") 
    # ******** NOTE THAT THE FOLDER STRUCTURE MUST MATCH WHAT IS ABOVE ^^^ EXACTLY.  **********
    # CHANGE out_dir AS NEEDED IF THERE ARE ANY DIFFERENCES IN THE LOCATION YOU WANT TO SAVE TO.
    
    if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  
    #df_group_icons_l file
    output_filename_group_icons <- paste0(project_name, "text_group_icons_HTML_formatted_", current_date, ".xlsx")
    write_xlsx(df_group_icons_l, file.path(out_dir, output_filename_group_icons)) #create file and save to 3ViewerPackages folder
    message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename_group_icons))
    
    #results_Top5 file
    output_filename_top5 <- paste0(project_name, "_top5_HTML_formatted", current_date, ".xlsx")
    write_xlsx(results_Top5, file.path(out_dir, output_filename_top5)) #create file and save to 3ViewerPackages folder
    message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename_top5))
    
    #text_installation (we call it df_bio)
    output_filename_inst_text <- paste0(project_name, "_text_installation_HTML_formatted_", current_date, ".xlsx")
    write_xlsx(df_bio, file.path(out_dir, output_filename_inst_text)) #create file and save to 3ViewerPackages folder
    message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename_inst_text)
    
    #text_group_desc
    output_filename_group_desc <- paste0(project_name, "_group_desc_HTML_formatted_", current_date, ".xlsx")
    write_xlsx(group_desc, file.path(out_dir, output_filename_inst_text)) #create file and save to 3ViewerPackages folder
    message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename_group_desc)
    
    
  ##create shortcut to Word to HTML folder ----
  
    #MT - this is still to do
  
    #bioclimatics shortcut
    out_full_path_bio <- file.path(out_dir, output_filename_bio) #save the path to the excel in 3ViewerPackages
    output_filelink_bio <- paste0(project_name, "_Bioclimatics_HTML_formatted_", current_date, ".lnk") #create shortcut name
    shortcut_location_bio <- file.path(input_dir, output_filelink_bio) #save the path to the future shortcut
    
    shell(paste0( #create shortcut to Word to HTML Conversion folder (this uses the Windows power shell)
      'powershell -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; ',
      '$s = $ws.CreateShortcut(\'', shortcut_location_bio, '\'); ',
      '$s.TargetPath = \'', out_full_path_bio, '\'; ',
      '$s.Save()"'
    )) 
    
# clean environment so that things can run properly for the next run  
#rm(list = ls()) 

