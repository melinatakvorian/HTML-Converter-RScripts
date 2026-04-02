# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2025.11.14


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

packages <- c("pandoc","xml2","rvest","writexl")

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
input_umbrella <- 
  "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/" 
#the specific folder inside the Document to HTML Table Converter where the input files are
input_specific_folder <- "_AirForceClimateViewerDev/Document to HTML Table Converter/FilesForTesting/Hydro_test/New_structure" 

#the final file name will start with this and will get the date added
project_name <- "NoSubheadings_VegStructure" #Replace with whatever you want.

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
  hist_indices <- c(1:14, last) # histclimatic sections
  disr_indices <- c(1:2, 15:(last-1)) # disruptions sections
  
  
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
    }else{df_hist[i, col] <- NA} #ChatGPT help
  }
}

##DISRUPTION SCENARIOS----

#find the indices within the list that are new occurrences of 'disruptions Group Name'
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
      
      mylist[[file]][[length(mylist[[file]])+1]] <- num_pair #create a nested list with each index within a disr group section
      
    }else{ #case for the last instance of new disr group
      num_pair <- c(num_files[[file]][[i]]:(length(results_disr[[file]])-1))
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
empty_cols <- c()

for(i in 1:ncol(test)){
  if(all(is.na(test[[i]]))){
    empty_cols[length(empty_cols)+1] <- i
  }
}

df_disr <- df_disr[ , -empty_cols]


##references hanging indent ----
#add REFERENCES SECTION HANGING INDENT <p style=???padding-left:15px;text-indent:-15px;???> 
for(i in 1:nrow(df_hist)){
  df_hist$References[i]
  #replace each <p> to <p style=???padding-left:15px;text-indent:-15px;???>
  temp_string <- df_hist$References[i]
  temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style="padding-left:15px;text-indent:-15px;">')
  df_hist$References[i] <- temp_string1
}

# Export final files ----
out_dir <- input_dir
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

#historical file
output_filename <- paste0(project_name, "_historical_HTML_formatted_", current_date, ".xlsx")
write_xlsx(df_hist, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

#disruptions group file
output_filename <- paste0(project_name, "_disruptions_HTML_formatted", current_date, ".xlsx")
write_xlsx(df_disr, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment so that things can run properly for the next run  
#rm(list = ls()) 

