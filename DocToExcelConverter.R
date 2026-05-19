# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2026.05.15


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

packages <- c("pandoc","xml2","rvest","writexl", "stringr", "readxl")

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
    input_umbrella <- "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/" 
    
    #NAVY
    #input_umbrella <- "N:/RStor/CEMML/ClimateChange/2_NavyClimate/Round2_Extremes_INRMP_integ/MidLant Region/"

    #the specific folder inside the Document to HTML Table Converter where the input files are
    input_specific_folder <- "Eielson AFB/TEVA/Word to HTML Conversion"
  
  #the final file name will start with this and will get the date added
    subject <- "Wildlife"
    installation <- "Eielson" #corresponds to shortName on the installation_info.xlsx
    project_name <- paste0(subject, "_", installation) #Replace with whatever you want.

#####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

  input_dir <-  paste0(input_umbrella, input_specific_folder) #Outputs will appear here as well.
  current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"
  installation_info <- readxl::read_xlsx("N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/_AirForceClimateViewerDev/Document to HTML Table Converter/FilesForTesting/Installation_Info.xlsx")

#ERROR CATCH 1 open files ----
  
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

convert_docx_to_html_full <- function(docx_file) {
  html_file <- tempfile(fileext = ".html")
  
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
  
  # ----- * remove '\r\n' from heading names -----
  #if sections[i] includes '\r\n', remove it
  clean_headings <- function(result_list){
    
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
results <- clean_headings(results)

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
  
# add full SITENAME, SITEID ----
  key <- match(installation, installation_info$ShortName)
  
  if(!is.na(key)){
    df[,"SITENAME"] <- installation_info$SITENAME[key]
    df[,"SITEID"] <- installation_info$SITEID[key]
  }else(print("No match found in installation database"))

  

##references hanging indent ----
#add REFERENCES SECTION HANGING INDENT <p style=???padding-left:15px;text-indent:-15px;???> 
for(i in 1:nrow(df)){
  df$References[i]
  #replace each <p> to <p style=padding-left:15px;text-indent:-15px;>
  temp_string <- df$References[i]
  temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style=padding-left:15px;text-indent:-15px;>')
  #temp_string2 <- replace_all_except_last(temp_string1, "</p>", "</p> <br>")
  df$References[i] <- temp_string1 #change to temp_string2 if you are adding the line breaks
}
  
##line breaks [manually input columns] ----
numbblocks <- c(3:20) # Change to the columns we need breaks in
    #add blank line after each paragraph
    for(a in 1:length(numbblocks)){
      col_num <- numbblocks[[a]]
      for(b in 1:nrow(df)){
        if(is.na(df[[col_num]][b])) next

        #replace each </p> to </p> <br>
        temp_string <- df[[col_num]][b]
        temp_string1 <- replace_all_except_last(temp_string, "</p>", "</p> <br>")
        df[[col_num]][b] <- temp_string1
      }
    }
 
  #for Hydro Qualitative conversion 
  # for(i in 1:nrow(df)){
  #   df$Installation_Summary[i]
  #   temp_string <- df$Installation_Summary[i]
  #   temp_string1 <- stringr::str_replace_all(temp_string, "</p> <ul>", "</p> <ul> <br>")
  #   temp_string2 <- replace_all_except_last(temp_string1, "</p>", "</p> <br>")
  #   df$Installation_Summary[i] <- temp_string2 #change to temp_string2 if you are adding the line breaks
  # }

# Export final files ----
out_dir <- input_dir
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

output_filename <- paste0(project_name, "_HTML_formatted_", current_date, ".xlsx")
write_xlsx(df, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment, so that things can run properly for the next run  
  #rm(list = ls()) 

