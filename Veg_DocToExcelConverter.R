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
      input_specific_folder <- "_AirForceClimateViewerDev/Document to HTML Table Converter/FilesForTesting/VegBMGR_test" 
      
  #the final file name will start with this and will get the date added
    project_name <- "refresh_run2" #Replace with whatever you want.
    
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
  parse_html_sections <- function(html_doc, section_indices) {
    #identify all headings
    headings <- rvest::html_nodes(html_doc, "h1") #identify headings
    sections <- vector("list", length(section_indices)) #create list of headings (sections)
    
    for (i in seq_along(section_indices)) { # Iterate over specified sections
        print(paste("Parsing section:", section_indices[i]))
        
      start_node <- headings[[section_indices[i]]]
      
      end_node <- if (i < length(section_indices)) headings[[section_indices[i + 1]]] else NULL #this works
        print(headings[section_indices[i + 1]])

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
      bio_indices <- c(1:5, last) # Bioclimatic sections
      veg_indices <- c(1:3, 6:(last - 1)) # Vegetation sections
      
      
      #Create BIO table list
      sections_bio <- parse_html_sections(html_doc, bio_indices)
      sections_bio <- sections_bio[names(sections_bio) != ""] #remove accidental headers
      results_bio[[basename(file)]] <- sections_bio #should be a list of headings and its text
      
      #Create VEG table listkjo
      sections_veg <- parse_html_sections(html_doc, veg_indices)
      sections_veg <- sections_veg[names(sections_veg) != ""] #remove accidental headers
      results_veg[[basename(file)]] <- sections_veg #should be a list of headings and its text
    }

#remove blank spaces after headings that could cause additional headers accidentally
  results_bio <- remove_end_blanks(results_bio)
  results_veg <- remove_end_blanks(results_veg)


#unfold the results list to be able to create a dataframe
  all_headings_bio <- unique(unlist(lapply(results_bio, names)))
  all_headings_veg <- c("SITENAME", "SITEID", "INRMPNAME")
  all_headings_veg <- append(all_headings_veg, unique(unlist(lapply(results_veg, names))))

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
      num_files <- as.list(c(1:as.numeric(length(results_veg))))
      
      #create mini lists for each instanve of new veg group
        for(file in seq_along(results_veg)){
          veg_names <- names(results_veg[[file]])
          indices <- c()
          
          for(i in seq_along(veg_names)){
            if(veg_names[i] =="New_Vegetation_Group") indices[length(indices)+1] <- i
          }
          num_files[[file]] <- indices
        }

    #use the indices to create smaller lists as keys to sections of Veg Groups in the document
      #initialize objects
      total_rows <- 0
      mylist <- vector("list", length(results_veg))
  
      #create mini lists, assign data to them
      for(file in seq_along(results_veg)){
        
        for(i in seq_along(num_files[[file]])){
          finish <- i+1
          
          if(finish <= length(num_files[[file]])){
            secondtolast <- num_files[[file]][[finish]]
            secondtolast <- secondtolast-1
            
            num_pair <- c(num_files[[file]][[i]]:secondtolast) #create range from one to the next
            
            total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
            
            mylist[[file]][[length(mylist[[file]])+1]] <- num_pair
            
          }else{
            num_pair <- c(num_files[[file]][[i]]:(length(results_veg[[file]])-1))
            total_rows <- total_rows + length(num_pair) #sum all iterations to see how long the df should be
            
            mylist[[file]][[length(mylist[[file]])+1]] <- num_pair
          }
        }
      }

    #create a df where each row is one of these lists. 
      df_veg <- data.frame(matrix(NA_character_, nrow=length(total_rows), ncol=length(unique(all_headings_veg))),
                           stringsAsFactors = FALSE)
      colnames(df_veg) <- unique(all_headings_veg)
      rownum <- 1
      
      for(file in seq_along(results_veg)){
        for(a in seq_along(mylist[[file]])){
          # Extract the current list of indices from mylist
          templist <- mylist[[file]][[a]]
          
          # Populate the first few columns with results_bio data (assuming it applies to all rows for this file)
          df_veg[rownum, 1] <- results_bio[[file]][[1]]
          df_veg[rownum, 2] <- results_bio[[file]][[2]]
          df_veg[rownum, 3] <- results_bio[[file]][[3]]
          
          n_col <- 4 # Start filling from the 4th column
          
          # Extract elements from results_veg based on the indices in templist
          for(b in seq_along(templist)){
            df_veg[rownum, n_col] <- results_veg[[file]][[templist[[b]]]]
            n_col <- n_col + 1
          }
          
          # Move to the next row for the dataframe
          rownum <- rownum + 1
        }
      }

  ##Delete empty columns ----
    test <- df_veg
    empty_cols <- c()
      
    for(i in 1:ncol(test)){
      if(all(is.na(test[[i]]))){
        empty_cols[length(empty_cols)+1] <- i
      }
    }
      
    df_veg <- df_veg[ , -empty_cols]

#make Exposure Icon column for Anthony
df_veg[, 'Exposure_Icon'] <- "Extreme Heat, Drought, Vector Borne Disease, Invasive Species, Seasonal Timing, Fire/Flooding"

  ##references hanging indent ----
    #add REFERENCES SECTION HANGING INDENT <p style=???padding-left:15px;text-indent:-15px;???> 
    for(i in 1:nrow(df_bio)){
      df_bio$References[i]
      #replace each <p> to <p style=???padding-left:15px;text-indent:-15px;???>
      temp_string <- df_bio$References[i]
      temp_string1 <- stringr::str_replace_all(temp_string, "<p>", '<p style="padding-left:15px;text-indent:-15px;">')
      df_bio$References[i] <- temp_string1
    }

# Export final files ----
  out_dir <- input_dir
  if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
  
  #bioclimatics file
  output_filename <- paste0(project_name, "_Bioclimatics_HTML_formatted_", current_date, ".xlsx")
  write_xlsx(df_bio, file.path(out_dir, output_filename))
  message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))
  
  #vegetation group file
  output_filename <- paste0(project_name, "_Vegetation_HTML_formatted", current_date, ".xlsx")
  write_xlsx(df_veg, file.path(out_dir, output_filename))
  message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# clean environment so that things can run properly for the next run  
#rm(list = ls()) 

