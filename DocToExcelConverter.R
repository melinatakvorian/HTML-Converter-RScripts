# Dashboard Word to Excel Converter v1.0
# CSU/CEMML - Trevor Lee Even, Ph.D.; Melina Takvorian, melina.takvorian@colostate.edu
# Date: 2025.11.14


# Converts .docx files in input_dir into dashboard-ready xlsx files.
# All headings must match across the document set. Paragraphs must be broken by a double carriage return. 
# Change project_name to an appropriate label for each dataset processed.
# All word documents in the target folder will be converted, so make sure you only have what you want in there.

# Set up ----

## Install / load necessary packages ----

  packages <- c("pandoc","xml2","rvest","writexl", "stringr", "officer","devtools", "RDCOMClient")

# Install packages not yet installed
  installed_packages <- packages %in% rownames(installed.packages())
  if (any(installed_packages == FALSE)) {
    install.packages(packages[!installed_packages])
  }

# load packages
  invisible(lapply(packages, library, character.only = TRUE))

## Create paths for storing files ----

# CHANGE AS DIRECTED BELOW ------------------------------------------------

#!!!PAY ATTENTION TO THE DIRECTION OF THE SLASHES. THEY HAVE TO BE CHANGED TO FORWARD SLASHES, AS SHOWN BELOW
  input_umbrella <- 
    "N:/RStor/CEMML/ClimateChange/1_USAFClimate/1_USAF_Natural_Resources/20_2_0004_RevisitingPhase1/_AirForceClimateViewerDev/" #broad folder structure
  input_specific_folder <- "Document to HTML Table Converter/FilesForTesting/TEVA_test" #the specific folder inside the Document to HTML Table Converter where the input files are
  
  input_dir <-  paste0(input_umbrella, input_specific_folder) #Rename to your target directory. Outputs will appear here as well.
  project_name <- "RDCOM_refresh1" #Replace with whatever you want.

# NO MORE CHANGES ------------------------------------------------------------------

  current_date <- format(Sys.Date(), "%Y%m%d")  # e.g., "2025-09-24"


# ----- * Word->HTML function with pandoc ----
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

# ---- * Word -> HTML function with RDCOMClient ----

  convert_rdccom <- function(docx_file, filepath){
    
    #html_file <- tempfile(fileext = ".html")
    html_file <- paste0(filepath, "/output", substr(docx_file, 186, 195), ".html") #place to store HTML before it gets read in
    
    wordApp <- COMCreate("Word.Application")
    wordApp[["Visible"]] <- FALSE #making the word application visible on the screen
    wordApp[["DisplayAlerts"]] <- FALSE
    #path_To_Word_File <- "C:/Users/melinata/OneDrive - Colostate/Desktop/TestFiles/TravisAFB_WildlandFire_Tester.docx"
    doc <- wordApp[["Documents"]]$Open(normalizePath(docx_file), ConfirmConversions = FALSE)
    wordApp[["ActiveDocument"]]$SaveAs2(
      FileName = normalizePath(html_file),
      FileFormat = 8
    )
    doc$close(FALSE)
    wordApp$Quit()
    
    xml2::read_html(html_file)
  }

  #example code
    # wordApp <- COMCreate("Word.Application")
    # wordApp[["Visible"]] <- TRUE
    # wordApp[["DisplayAlerts"]] <- FALSE
    # path_To_Word_File <- "C:/Users/melinata/OneDrive - Colostate/Desktop/TestFiles/TestRDCOM/TravisAFB_WildlandFire_Tester.docx"
    # doc <- wordApp[["Documents"]]$Open(normalizePath(docx_files[[1]]), ConfirmConversions = FALSE)
    # wordApp[["ActiveDocument"]]$SaveAs2(
    #   FileName = normalizePath("C:/Users/melinata/OneDrive - Colostate/Desktop/TestFiles/TestRDCOM/output.html"),
    #   FileFormat = 8
    # )
    # doc$close(FALSE)
    # wordApp$Quit()

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

# ----- * #removing spaces after headings function -----
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
    #html_doc <- convert_docx_to_html_full(file) #old
    html_doc <- convert_rdccom(file, input_dir) #new
    sections <- parse_html_sections(html_doc)
    sections <- sections[names(sections) != ""] #remove accidental headers
    results[[basename(file)]] <- sections
   }

#remove trailing spaces from headings
  results <- remove_end_blanks(results)

#unfold the results list to be able to create a dataframe
  all_headings <- unique(unlist(lapply(results, names))) #THIS SHOULD BE 37, IF THE LOOP ABOVE WORKED


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

#Check for incorrect naming of the installation ID, correct it
  colnames(df)[colnames(df) == "Installation"] <- "SITENAME"
  colnames(df)[colnames(df) == "Site_Name"] <- "SITENAME"

# replace Microsoft stylization issues ----
  #replace the <o:p> or </o:p> with <p> </p>
  for(i in seq_along(df)){ #for each column
    for(cell in seq_along(df[[i]])){ #for each cell in each column
      print(df[[i]][cell])
      df[[i]][cell] <- stringr::str_replace_all(
        df[[i]][cell], c("<o:p>" = "<p>", 
                         "</o:p>" = "</p>", 
                         ' class=\\"MsoNormal\\"' = "",
                         "<span style=\"mso-spacerun:yes\"> </span>" = ""))# add line for MS office formatting
      print(df[[i]][cell])
    }
  }

# Export final files ----
out_dir <- input_dir
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

output_filename <- paste0(project_name, "_HTML_formatted_", current_date, ".xlsx")
write_xlsx(df, file.path(out_dir, output_filename))
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment, so that things can run properly for the next run  
#rm(list = ls()) 


