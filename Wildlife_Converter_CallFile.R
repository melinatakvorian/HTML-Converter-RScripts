# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2026.05.15


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

  ## Install / load necessary packages ----
    
    packages <- c("pandoc","xml2","rvest","writexl", "stringr", "readxl", "dplyr")
    
    # Install packages not yet installed
    installed_packages <- packages %in% rownames(installed.packages())
    if (any(installed_packages == FALSE)) {
      install.packages(packages[!installed_packages])
    }
    
    # load packages
    invisible(lapply(packages, library, character.only = TRUE))
  
  ## Create paths for storing files ----

    #####CHANGE AS DIRECTED BELOW --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---
    
    #PAY ATTENTION TO THE DIRECTION OF THE SLASHES. THEY HAVE TO BE CHANGED TO FORWARD SLASHES, AS SHOWN BELOW
    #the broad folder structure
    
    #AIR FORCE  
      #input_umbrella <- "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/"
      
      input_umbrella <- "N:/RStor/CEMML/ClimateChange/0_Natural Resources Teams/Wildlife/_FWVAs/" #being used for testing
    
    #NAVY
      #input_umbrella <- "N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/MidLant Region/"
    
    #the specific folder inside the Document to HTML Table Converter where the input files are
      input_installation_folder <- "Air Force"
      installation_type <- "Air Force" #"Navy"
      input_SME_folder <- "/Travis"
    
    #the final file name will start with this and will get the date added
      subject <- "FWVA"
      project_name <- paste0(subject, "_", input_installation_folder) 
    
    #####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

    input_dir <-  paste0(input_umbrella, input_installation_folder, input_SME_folder)
    current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"


# ERROR CATCH: open files ----

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

# Run the Function Library ----
  source("FunctionLibrary.R", echo = TRUE)
  
# RUN ----
  docx_files <- list.files(input_dir, pattern = "\\.docx$", full.names = TRUE) #pull list of all files in folder
  docx_files <- docx_files[!grepl("^~\\$", basename(docx_files))]
  
  results <- list()
  
  for (file in docx_files) { #for each file, convert it to HTML, Identify its sections, delete empty headers, add to a results mega-list
    html_doc <- convert_docx_to_html_full(file)
    sections <- parse_html_sections(html_doc)
    sections <- sections[names(sections) != ""] #remove accidental headers
    results[[basename(file)]] <- sections
  }

  #QAQC heading names for trailing spaces and line breaks
  results <- remove_end_blanks(results)
  results <- remove_accidental_return(results)
  
  #unfold the results list to be able to create a dataframe
  all_headings <- unique(unlist(lapply(results, names)))


# Create dataframe and input HTML in proper sections ----
  df <- data.frame(matrix(NA_character_, length(results), length(all_headings)),
                   stringsAsFactors = FALSE)
  colnames(df) <- all_headings
  rownames(df) <- names(results)
  for (i in seq_along(results)) {
    for (col in names(results[[i]])) {
      df[i, col] <- results[[i]][[col]]
    }
  }

# run paragraph notation editor ----
  if(subject == "TEVA"){
    #TEVAs
    cols_to_change <- c("SITEID", "CommonName", "ScientificName", "SpeciesIDNum", "FedTxt",
                        "StateTxt", "AdditionalStatus", "Presence", "BreedingStatus",
                        "FirstHabitat", "SecondHabitat", "ThirdHabitat", "FourthHabitat",
                        "VulnerabilityResult", "Confidence", "VulnSummary", "NE_Text", "NE_Level", "OE_Level",
                        "OE_Text", "S_Text", "S_Level", "AC_Text", "AC_Level", "ReferencesTxt")
  }else if(subject == "FWVA"){
    #FWVAs
    cols_to_change <- c("SITEID","HabitatCommunity", "HabitatCommIDNum",
                        "FirstHabitat", "SecondHabitat", "ThirdHabitat", "FourthHabitat",
                        "VulnerabilityResult", "E_Text", "E_Level", "S_Text",
                        "S_Level", "AC_Text", "AC_Level")
  }
  
  df <- p_be_gone(df, cols_to_change)

# remove italics from sci names ----
  if(subject == "TEVA"){
    df$ScientificName <- stringr::str_replace_all(df$ScientificName, "<em>", '')
    df$ScientificName <- stringr::str_replace_all(df$ScientificName, "</em>", '')
  }else(
    print("Not a TEVA run - Scientific Names are N/A. Nothing will be changed at this step.")
  )


# add full SITENAME, SITEID ----
  if(installation_type == "Navy"){
    for(i in 1:nrow(df)){
      installation_info <- readxl::read_xlsx("N:/RStor/CEMML/ClimateChange/Document Standards/Templates/TEMPLATES_SME_Word_Docs/Installation_IDs.xlsx", sheet=2)
      
      #create SITENAME and assign the value from the corresponding row of the excel spreadsheet according to SITEID
      SITENAME <- installation_info$InstallationNames[installation_info$SITEID == df$SITEID[i]] 
      
      #assign the correct InstallationNames to that row of data
      df[i,"InstallationNames"] <- SITENAME
      
      #move the InstallationNames to the correct row 
      df <- df %>% relocate(InstallationNames, .after = SITEID)
      
      #create InstallationID and assign the value from the corresponding row of the excel spreadsheet according to SITEID
      InstallationID <- installation_info$`Installation ID (Site Code)`[installation_info$SITEID == df$SITEID[i]]
      
      #assign the correct InstallationNames to that row of data
      df[i,"Installation ID (Site Code)"] <- InstallationID
      
      #move the new column after the SITEID column
      df <- df %>% relocate(`Installation ID (Site Code)`, .after = SITEID)
    }
  }else if(installation_type == "Air Force"){
    for(i in 1:nrow(df)){
      installation_info <- readxl::read_xlsx("Installation_IDs.xlsx", sheet=1)
      
      #create SITENAME and assign the value from the corresponding row of the excel spreadsheet according to SITEID
      SITENAME <- installation_info$SITENAME[installation_info$SITEID == df$SITEID[i]]
      
      #assign the correct SITENAME to that row of data
      df[i,"SITENAME"] <- SITENAME
      
      #move the SITENAME to the correct row 
      df <- df %>% relocate(SITENAME, .after = SITEID)
    }
  }

#references hanging indent ----
  df <- ref_hanging_indents(df, subject)

#change US type ----
  df <- update_US(df, subject, installation_type)

# create hex codes and numbers ----
  df <- hex_codes(df, subject)

# add habitat_icons column ----
  df <- habitat_icons(df)

#color the word with the vulnerability score ---
  df <- color_vuln_text(df)

#ADD THE BOLDING FUNCTION CODE
  
# Export final files ----
  ##export excel to 3ViewerPackages folder ----
    out_dir <- paste0(input_umbrella, input_installation_folder, "/3ViewerPackages/HTML_excels") 
    
    # ******** NOTE THAT THE FOLDER STRUCTURE MUST MATCH WHAT IS ABOVE ^^^ EXACTLY.  **********
    # CHANGE out_dir AS NEEDED IF THERE ARE ANY DIFFERENCES IN THE LOCATION YOU WANT TO SAVE TO.
    
    if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
    
    output_filename <- paste0(project_name, "_HTML_formatted_", current_date, ".xlsx")
    #write_xlsx(df, file.path(out_dir, output_filename)) #create file and save to 3ViewerPackages folder
    write_xlsx(result, file.path(out_dir, output_filename)) #create file and save to 3ViewerPackages folder
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


# clean environment, so that things can run properly for the next run  
#rm(list = ls()) 

