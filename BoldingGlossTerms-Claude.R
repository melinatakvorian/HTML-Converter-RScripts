# ---------------------------------------------------------------
# bold_first_term_occurrence()
#
# For each row of `df`, looks across the columns in `cols` (in the
# order you give them) and, for every term in `glossary`, wraps ONLY
# the first occurrence of that term (across all selected columns)
# in <strong></strong>. Writes the result back into the same cells.
#
# - Longer/multi-word terms are matched before shorter ones, so
#   "heart attack" claims its match before "heart" does.
# - Matching uses word boundaries (\\b) so "art" won't match inside
#   "heartache".
# - Matches inside text you've already bolded (from an earlier,
#   longer term) are skipped, so you never get nested/broken tags.
# ---------------------------------------------------------------

escape_regex <- function(x) {
  # Escape one special character at a time using fixed = TRUE, so no
  # regex engine ever has to parse a pattern containing these
  # characters together (which is what caused the invalid regex error).
  # IMPORTANT: backslash must be escaped first, or it would double-escape
  # the backslashes inserted for the other characters.
  specials <- c("\\", ".", "|", "(", ")", "[", "]", "{", "}", "^", "$", "*", "+", "?")
  for (ch in specials) {
    x <- gsub(ch, paste0("\\", ch), x, fixed = TRUE)
  }
  x
}

bold_first_term_occurrence <- function(df, cols, glossary, ignore_case = TRUE) {
  
  # longest terms first, so multi-word phrases win over their substrings
  glossary <- glossary[order(-nchar(glossary))]
  
  for (i in seq_len(nrow(df))) {
    
    texts <- setNames(as.character(unlist(df[i, cols])), cols)
    
    # track character ranges in each column that are already inside a
    # <strong> tag, so later terms don't re-match inside them
    protected <- setNames(vector("list", length(cols)), cols)
    for (col in cols) protected[[col]] <- matrix(numeric(0), ncol = 2)
    
    for (term in glossary) {
      pattern <- paste0("\\b", escape_regex(term), "\\b")
      found <- FALSE
      
      for (col in cols) {
        if (found) break
        txt <- texts[[col]]
        if (is.na(txt) || !nzchar(txt)) next
        
        m <- gregexpr(pattern, txt, ignore.case = ignore_case, perl = TRUE)[[1]]
        if (m[1] == -1) next
        lens <- attr(m, "match.length")
        prot <- protected[[col]]
        
        for (k in seq_along(m)) {
          start <- m[k]; end <- start + lens[k] - 1
          
          overlap <- FALSE
          if (nrow(prot) > 0) {
            overlap <- any(start <= prot[, 2] & end >= prot[, 1])
          }
          if (overlap) next
          
          matched_text <- substr(txt, start, end)
          new_txt <- paste0(
            substr(txt, 1, start - 1),
            "<strong>", matched_text, "</strong>",
            substr(txt, end + 1, nchar(txt))
          )
          
          # shift any already-protected ranges that came after this match
          tag_len <- nchar("<strong>") + nchar("</strong>")
          if (nrow(prot) > 0) {
            shift <- ifelse(prot[, 1] > end, tag_len, 0)
            prot[, 1] <- prot[, 1] + shift
            prot[, 2] <- prot[, 2] + shift
          }
          new_start <- start + nchar("<strong>")
          new_end   <- new_start + (end - start)
          prot <- rbind(prot, c(new_start, new_end))
          
          protected[[col]] <- prot
          texts[[col]] <- new_txt
          found <- TRUE
          break
        }
      }
    }
    
    df[i, cols] <- as.list(texts[cols])
  }
  
  df
}

#TEST ----
glossary <- c("weather", "vulnerability", "natural hazards", "temperature-dependent sex determination", "drought", "frequency", "duration")
sections <- c("VulnSummary", "NE_Text", "OE_Text", "S_Text", "AC_Text")

result <- bold_first_term_occurrence(df, cols = sections, glossary = glossary)
print(result)


#EXPORT ----
##export excel to 3ViewerPackages folder ----
out_dir <- paste0(input_umbrella, input_installation_folder, "/3ViewerPackages/HTML_excels") 

# ******** NOTE THAT THE FOLDER STRUCTURE MUST MATCH WHAT IS ABOVE ^^^ EXACTLY.  **********
# CHANGE out_dir AS NEEDED IF THERE ARE ANY DIFFERENCES IN THE LOCATION YOU WANT TO SAVE TO.

if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

output_filename <- paste0(project_name, "_HTML_formatted.xlsx")
shortcut_location <- file.path(input_dir, output_filename) #save the path to the future shortcut

write_xlsx(result, shortcut_location) #create file and save to folder
message("Conversion complete. XLSX saved to: ", file.path(out_dir, output_filename))

# ---------------------------------------------------------------
# Example usage ----
df <- data.frame(
  colA = c("Patient had a heart attack.", "No cardiac issues noted."),
  colB = c("The heart attack occurred at home.", "Follow-up for heart attack risk."),
  stringsAsFactors = FALSE
)

glossary <- c("heart attack", "heart")

result <- bold_first_term_occurrence(df, cols = c("colA", "colB"), glossary = glossary)
print(result)
#
# Expected for row 1:
#  - "heart attack" (as a phrase) gets its first occurrence bolded in
#    colA: "Patient had a <strong>heart attack</strong>."
#  - "heart" (as its own glossary term) is NOT re-bolded inside that
#    same "heart attack" span (it's protected), so its own first
#    occurrence is found next, in colB:
#    "The <strong>heart</strong> attack occurred at home."
#  Each glossary term is tracked independently -- "heart" and "heart
#  attack" both get exactly one bolded occurrence each, and neither
#  steals territory already claimed by the other.