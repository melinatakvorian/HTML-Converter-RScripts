# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2025.12.10


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

packages <- c("pandoc","xml2","rvest","writexl", "dplyr")

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
      input_umbrella <- 
      "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/" 
    #the specific folder inside the Document to HTML Table Converter where the input files are
      input_specific_folder <- "Dover AFB/3ViewerPackages/Ecosystems/Word to HTML convertion" 
  
  #the final file name will start with this and will get the date added
    project_name <- "Dover_run1_Veg" #Replace with whatever you want.

#####NO MORE CHANGES --- -- -- -- --- - - -- -- - -  - - - - -  --- - - - - - - --- --- --- -- ---

  input_dir <-  paste0(input_umbrella, input_specific_folder) #Rename to your target directory. Outputs will appear here as well.
  current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"


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
  # Identify all h1 headings
  h1_headings <- rvest::html_nodes(html_doc, "h1")
  sections <- vector("list", length(h1_headings)) # Create list for h1 sections
  
  for (i in seq_along(h1_headings)) {
    # Start node for the current h1 heading
    start_node <- h1_headings[[i]]
    # End node for the next h1 heading (if it exists)
    end_node <- if (i < length(h1_headings)) h1_headings[[i + 1]] else NULL
    
    # Identify all siblings following the h1 heading
    #siblings <- xml2::xml_find_all(start_node, "following-sibling::*")#ORIGINAL
    siblings <- xml2::xml_find_all(start_node, "following-sibling::*[self::h1 or self::h2]")
    if (!is.null(end_node)) {
      idx <- which(vapply(siblings, identical, logical(1), y = end_node))
      if (length(idx) == 0) idx <- length(siblings) + 1
      siblings <- siblings[seq_len(idx - 1)]
    }
    
    # Filter the siblings to identify h2 subheadings
    h2_subheadings <- siblings[xml2::xml_name(siblings) == "h2"]
    
    # Create a named list for h2 subheadings
    sub_sections <- vector("list", length(h2_subheadings))
    for (j in seq_along(h2_subheadings)) {
      h2_start_node <- h2_subheadings[[j]]
      h2_end_node <- if (j < length(h2_subheadings)) h2_subheadings[[j + 1]] else NULL
      
      h2_siblings <- xml2::xml_find_all(h2_start_node, "following-sibling::*")#ORIGINAL
      # #h2_siblings <- xml2::xml_find_all(h2_start_node, "following-sibling::h2")
      # if (!is.null(h2_end_node)) {
      #   idx <- which(vapply(h2_siblings, identical, logical(1), y = h2_end_node))
      #   if (length(idx) == 0) idx <- length(h2_siblings) + 1
      #   h2_siblings <- h2_siblings[seq_len(idx - 1)]
      # } #ORIGINAL
      
      
      #CHATGPT SUGGESTION ---
      if (!is.null(h2_end_node)) {
        idx <- which(vapply(h2_siblings, identical, logical(1), y = h2_end_node))
        if (length(idx) == 0) idx <- length(h2_siblings) + 1
        h2_siblings <- h2_siblings[seq_len(idx - 1)]
      } else if (!is.null(end_node)) {
        idx <- which(vapply(h2_siblings, identical, logical(1), y = end_node))
        if (length(idx) == 0) idx <- length(h2_siblings) + 1
        h2_siblings <- h2_siblings[seq_len(idx - 1)]
      } else {
        h2_siblings <- h2_siblings
      }
      
      
      
      # Concatenate the content under each h2 subheading
      h2_content_html <- paste(as.character(h2_siblings), collapse = " ")
      # Directly assign the content to the sub_sections list
      sub_sections[[j]] <- h2_content_html
    }
    
    # Name the sub_sections list based on the h2 titles
    names(sub_sections) <- vapply(h2_subheadings, xml2::xml_text, "")
    
    # Concatenate the content under the h1 heading (excluding h2 subheadings)
    h1_content_html <- paste(as.character(siblings[!siblings %in% h2_subheadings]), collapse = " ")
    # Directly assign the content to the sections list
    sections[[i]] <- list(content = h1_content_html, sub_sections = sub_sections)
  }
  
  # Assign names to the sections list based on h1 titles
  names(sections) <- vapply(h1_headings, xml2::xml_text, "")
  sections
}

# ----- * REMOVING SPACES after headings function -----
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

#remove blank spaces after headings that could cause additional headers accidentally
results <- remove_end_blanks(results)

#unfold the results list to be able to create a dataframe
all_headings <- unique(unlist(lapply(results, names)))

# Make initial data frame ----
scenarios <- all_headings[! all_headings %in% c("SITEID", "SITENAME")]

table_cols <- c()

for(sec in 1:length(sections)){
  section_item <- sections[[sec]]
  for(subsec in 1:length(section_item$sub_sections)){ #identify the h1 section
    subsec_name <- names(section_item$sub_sections[subsec])
    
    if(!is.na(subsec_name)&&length(subsec_name)>0){ #make sure there actually is a subsection to pull from
      table_cols <- c(table_cols, subsec_name) #pull the h2 section name
    }
  }
}

table_cols <- unique(table_cols)
table_cols <- c(table_cols, c("SITENAME", "SITEID"))

df <- data.frame(matrix(NA_character_, length(scenarios), length(table_cols)),
                 stringsAsFactors = FALSE)

colnames(df) <- c(table_cols)
rownames(df) <- scenarios

df <- df %>% relocate(c(SITENAME, SITEID), .before = Scenario)

for (h1 in seq_along(results)) {
  # Check if results[[h1]] is valid
  if (!is.null(results[[h1]]) && length(results[[h1]]) > 0) {
    
    for (h2 in seq_along(results[[h1]])) {
      
      # Check if results[[h1]][[h2]] is valid
      if (!is.null(results[[h1]][[h2]]) && !is.null(results[[h1]][[h2]]$sub_sections)) {
        
        for (content in seq_along(results[[h1]][[h2]]$sub_sections)) {
          # Check if sub_sections[[content]] is valid
          text <- results[[h1]][[h2]]$sub_sections[[content]]
          
          if (!is.null(text) && !is.na(text) && length(text) > 0) {
            row <- h2 - 2
            if(startsWith(names(results[[h1]][[h2]]$sub_sections[2]), "Period: Historical") || content == 1){
              col <- content+2
            }else {
              col <- content+5
            }
            
            df[row, col] <- text
            
            df$SITENAME <- results[[h1]]$SITENAME$content
            df$SITEID <- results[[h1]]$SITEID$content
            
            print(paste("h1:", h1, "h2:", h2, "content:", content))
          }
        }
      }
    }
  }
}

# Format df according to Anthony's table ----
#we want a table with the same information, but a different structure

df <- df %>% relocate(c(
  `Period: Near Term, SPEI_Text`, 
  `Period: Far Term, SPEI_Text`, 
  `Period: Historical, SPEI_Text`,), .before = Dry_Distribution_Text)

SPEI_txt <- df %>% select(Scenario,
                          `Period: Near Term, SPEI_Text`,
                          `Period: Far Term, SPEI_Text`,
                          `Period: Historical, SPEI_Text`,
                          Installation_Summary)

rownames(SPEI_txt) <- SPEI_txt[,1] #assign row names, these will become the columns
SPEI_txt <- SPEI_txt %>% select(-Scenario)

SPEI_txt <- as.data.frame(t(SPEI_txt))

#REBUILD into final dataset

foundation <- df %>% select(SITENAME, SITEID, Scenario, References, `Period: Historical, SPEI_Text`)
foundation <- df
rownames(foundation) <- NULL

#assign period to values
foundation[ , "Period"] <- NA  
foundation <- foundation %>% relocate(Period, .after = Scenario)

for(scen in 1:nrow(foundation)){
  if(is.na(foundation$Scenario[scen])){ 
    next }
  if(foundation$Scenario[scen] == "<p>Historical</p>"){
    foundation$Period[scen] <- "<p>Historical</p>"
  }else if(foundation$Scenario[scen] == "<p>High Disruption</p>" || foundation$Scenario[scen] == "<p>Moderate Disruption</p>" ){
    end <- nrow(foundation)+1
    foundation[end,] <- foundation[scen,] #copy this line to the end of the table
    foundation$Period[scen] <- "<p>Near Term</p>"
  }
}

for(scen in 1:nrow(foundation)){
  if(is.na(foundation$Scenario[scen])){ 
    next }
  if(is.na(foundation$Period[scen])){
    foundation$Period[scen] <- "<p>Far Term</p>"
  }
}

#create SPEI_text column
foundation[, "SPEI_text"] <- NA
foundation <- foundation %>% relocate(SPEI_text, .after = Period)

for (i in seq_len(nrow(foundation))) {
  print(paste("Row:", i, "Period:", foundation$Period[i]))
  foundation$Installation_Summary[i] <- df$Installation_Summary[1]
  
  if(is.na(foundation$Scenario[i])){ 
    next 
  } else if (foundation$Period[i] == "<p>Far Term</p>") {
    foundation[i, 5] <- foundation$`Period: Far Term, SPEI_Text`[i]
  } else if (foundation$Period[i] == "<p>Near Term</p>") {
    foundation[i, 5] <- foundation$`Period: Near Term, SPEI_Text`[i]
  } else {
    foundation$SPEI_text[i] <- foundation$`Period: Historical, SPEI_Text`[i]
  }
}

formatted <- foundation[!is.na(foundation$Scenario),]
formatted <- formatted %>% select(-c(`Period: Near Term, SPEI_Text`, `Period: Far Term, SPEI_Text`, `Period: Historical, SPEI_Text`))

missing_cols <- c()

for(i in 1:ncol(formatted)){
  
  if(all(is.na(formatted[,i]))){
    missing_cols <- c(missing_cols, i)
  }
}

if(!is.null(missing_cols)){
  missing_cols <- as.numeric(missing_cols)
  formatted_check <- formatted[,-missing_cols]
}else(formatted_check <- formatted)

# Adding Indent notation to HTML code ----

#add REFERENCES SECTION HANGING INDENT <p style=???padding-left:15px;text-indent:-15px;???> 
indent_df1 <- formatted_check
for(i in 1:nrow(indent_df1)){
  indent_df1$References[i]
  #replace each <p> to <p style=???padding-left:15px;text-indent:-15px;???>
  temp_string <- indent_df1$References[i]
  temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style="padding-left:15px;text-indent:-15px;">')
  indent_df1$References[i] <- temp_string1
}


#add indentation to paragraphs in other sections
indent_df2 <- indent_df1

columns_to_indent <- c("SPEI_text", "Installation_Summary", "Dry_Distribution_Text", "Wet_Distribution_Text", 
                       "Dry_Duration_Severity_Text", "Wet_Duration_Severity_Text")

for (i in 1:length(columns_to_indent)) { # iterate through list of columns that need indented paragraphs
  for (col in 1:ncol(indent_df2)) { # iterate through table to identify the columns
    
    if (colnames(indent_df2)[col] == columns_to_indent[i]) { # identify matching column names
      
      for (row in 1:nrow(indent_df2)) { # iterate through the rows of each column to add indentation
        
        temp_string <- indent_df2[[col]][row]
        detect <- stringr::str_detect(temp_string, "</p> <p>")
        
        # First condition: Add indentation to <p> tags at the start of the string
        if (!is.na(temp_string) && startsWith(temp_string, "<p>") && substr(temp_string, 4, 4) != "<") { 
          temp_string <- stringr::str_replace(temp_string, "<p>", '<p style="text-indent:15px;">')
        } 
        
        # Second condition: Handle "</p> <p>" with indentation if the following character is NOT "<" or " "
        if (!is.na(temp_string) && detect) { # Simplified condition
          coords <- stringr::str_locate_all(temp_string, "</p> <p>")[[1]] # Extract matrix of match positions
          
          if (!is.null(coords)) { # Ensure matches exist
            for (j in 1:nrow(coords)) { # Iterate through matches
              start_pos <- coords[j, 1]
              end_pos <- coords[j, 2]
              
              # Check the character after "</p> <p>"
              next_char <- substr(temp_string, end_pos + 1, end_pos + 1)
              if (next_char != "<" && next_char != " ") {
                # Replace the current "</p> <p>" with the indented version
                temp_string <- stringr::str_sub(temp_string, 1, start_pos - 1) %>% 
                  paste0('</p> <p style="text-indent:15px;">', stringr::str_sub(temp_string, end_pos + 1))
              }
            }
            indent_df2[[col]][row] <- temp_string # Update the dataframe cell
          }
        }
      }
    } else {next}
  }
}



# Export final files ----
out_dir <- input_dir
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

output_filename <- paste0(project_name, "_HTML_formatted", current_date, ".xlsx")
write_xlsx(indent_df2, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment, so that things can run properly for the next run  
rm(list = ls())
