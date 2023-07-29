
#' Extracts the PMIDs from the word documents in the Dropbox folder
#'
#' @param wd.data Path to the folder in which the word documents, having the PMIDs, are.
#'
#' @return -
#' @export
#'
#' @examples -
#'

Extractor <- function(wd.data, callout=T){
  library(readtext)

  if(!exists("wd.data")){wd.data <- "C:/Users/TPreijers/Dropbox/4Abstracts/R_test"}

  fl.folder <- list.files(wd.data, include.dirs = T, pattern = "^4CP - Editie \\d{2,3} - \\d{4}.*$")
  if(!exists("fl.folder") | length(fl.folder)==0){stop("There is no 4CP edition folder from Dropbox!")}

  val.last <- tail(sort(as.numeric(gsub("^4CP - Editie (\\d{2,3}) - \\d{4}.*$", "\\1", fl.folder))), 1)
  ch.folder <- fl.folder[grepl(paste0("^4CP - Editie ", val.last, ".*$"), fl.folder)]

  fl.files <- dir(path = file.path(wd.data, ch.folder), pattern = "docx", full.names = T)

  pmids <- NA

  for(i in fl.files){
    doc.text  <- readtext(file = i)
    doc.parts <- strsplit(doc.text$text, "\n")[[1]]
    doc.pmids <- doc.parts[grepl("pmid", tolower(doc.parts))]

    if(length(grepl("PMID", doc.pmids))>0){
      doc.pmids[grepl("PMID", doc.pmids)] <- gsub("^.*?PMID: (\\d{8,9})>*?$", "\\1", doc.pmids[grepl("PMID", doc.pmids)])
    }

    pmids <- append(pmids, gsub(".*(\\d{8,}).*?","\\1",  doc.pmids))

    if(sum(grepl("DOI", pmids))>0){
      pmids[grepl("DOI", pmids)] <- gsub(".*?(\\d{8,9}) .*", "\\1", pmids[grepl("DOI", pmids)], )
    }
  }

  pmids <- c(pmids[2:length(pmids)])

  if(callout){print(pmids)}


  dat.val <- paste0(gsub("-", "", Sys.Date()), "_selectie.txt")

  writeLines(con=paste0(ch.folder, dat.val), text = pmids)
}
