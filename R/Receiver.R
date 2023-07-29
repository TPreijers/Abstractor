
#' Retrieves the PMIDs from the word documents from the Dropbox of 4CP
#'
#' @param wd.data Path to the folder in which the word documents, having the PMIDs, are.
#'
#' @return -
#' @export
#'
#' @examples -
#'


Receiver <- function(wd.data){

  library(easyPubMed)
  library(XML)
  library(xml2)
  library(officer)
  library(textutils)
  library(stringr)

  ## 1. Retrieve pubmed query (DUMMY)
  #pm_query <- '32500636[pmid] OR 31925806[pmid] OR 32128848[pmid] OR 32022291[pmid] OR 32434285[pmid] OR 32157630[pmid] OR 32940349[pmid] OR 32939689[pmid]'

  ## 1. load input file
  if(!exists("wd.data")){wd.data <- "C:/Users/TPreijers/Dropbox/4Abstracts/R_test"}

  lst.file <- list.files(wd.data, pattern = ".txt")
  dat.pmid <- read.delim(file.path(wd.data, lst.file[length(lst.file)]), header = F)
  dat.pmid <- data.frame(trimws(dat.pmid[,1]))

  ## Sort of sanity check
  if(sum(!grepl("^\\d{8}$", dat.pmid[,1]))>0){
    cat("Non-valid rows present in input-file.")
    cat("The following rows were deleted:\n")
    print(dat.pmid[!grepl("^\\d{8}$", dat.pmid[,1]),,drop=FALSE])
    stop()
  } else {
    cat("Input-file correct.\n")
  }

  dat.pmid <- dat.pmid[grepl("^\\d{8}$", dat.pmid[,1]),,drop=FALSE]


  ## 1. Retrieve pubmed query
  pm_query <- paste0(paste0(dat.pmid[,1], "[pmid]"), collapse = " OR ")

  pm_entrez_id     <- get_pubmed_ids(pm_query)
  pm_abstracts_xml <- fetch_pubmed_data(pubmed_id_list = pm_entrez_id)
  pm_abstracts_abs <- fetch_pubmed_data(pubmed_id_list = pm_entrez_id, format = "abstract")

  ## Remove empty data rows
  if(pm_abstracts_abs[1] %in% ""){pm_abstracts_abs <- pm_abstracts_abs[-1]}
  if(pm_abstracts_abs[1] %in% ""){pm_abstracts_abs <- pm_abstracts_abs[-1]}

  ## Grep lines with usefull info
  wh.line       <- which(grepl("^\\d{1,}\\..*doi:.*", pm_abstracts_abs))
  pm_select_abs <- pm_abstracts_abs[grepl("^\\d{1,}\\..*doi:.*", pm_abstracts_abs)]


  ## Grep only lines with data about published article
  for(i in wh.line){
    if(pm_abstracts_abs[i+1] == ""){next}
    if(pm_abstracts_abs[i+1] != ""){pm_select_abs[which(wh.line %in% i)] <- paste0(pm_select_abs[which(wh.line %in% i)], pm_abstracts_abs[i+1], collapse = " ")}
  }

  ## Fetch usefull data, retreived using regex. Ahead of print articles are excluded from data inclusion using an XXX flag
  pm.extract <- data.frame(
    jou = gsub("^\\d{1,2}\\.\\s(.*)\\s\\d{4}.*doi.*$", "\\1", pm_select_abs), #journal
    mon =  gsub(".*\\d{4}\\s([a-zA-Z]{3}).*doi.*", "\\1", pm_select_abs), # month
    iss = ifelse(grepl("ahead", pm_select_abs ), "EOP", gsub("^.*\\((\\d+)\\).*doi.*", "\\1", pm_select_abs)), # issue
    vol = ifelse(grepl("ahead", pm_select_abs ), "EOP", gsub("^.*;(\\d+)\\(\\d+\\).*doi.*", "\\1", pm_select_abs)), # volume
    pag = ifelse(grepl("ahead", pm_select_abs ), "EOP", gsub("^.*:(\\d+\\-\\d+)\\..*doi.*", "\\1", pm_select_abs)) # volume
  )


  ## Also retrieve PMID, for checking if order is similar between pm.table and pm.extract.
  pm.extract$pmid <- as.numeric(trimws(gsub("PMID: (\\d{7,9})(.*)?", "\\1", pm_abstracts_abs[which(grepl("PMID:.*", pm_abstracts_abs))])))


  ## Create months dataset, for change abbrev into full
  ds.months <- data.frame(mon = c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
                          rep = c("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"))


  for(i in 1:nrow(pm.extract)){
    pm.extract$mon[i] <- ds.months$rep[ds.months$mon %in% pm.extract$mon[i]]
  }


  ## create extra dataset for dataextraction based on regex
  abstr.extract <- pm_abstracts_xml


  ## 2. Start creating datatable for retrieving info, beware not all data is in the pm_abstracts_xml datatable
  pm.table <- table_articles_byAuth(pm_abstracts_xml, max_chars = 5000)

  pm.table <- pm.table[,c("pmid","doi","title","abstract","year","month","lastname","firstname","jabbrv")]

  pm_abstracts_xml <- xmlParse(pm_abstracts_xml)


  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Abstract"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Volume"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//MedlinePgn"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Issue"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Month"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Year"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Title"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//PubDate"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml, "//AuthorList//LastName"))
  # xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml, "//AuthorList//ForeName"))

  wh.sel.pmd <- pm.extract$pmid[!grepl("EOP", pm.extract$iss)]
  wh.sel.abs <- pm.table$pmid %in% wh.sel.pmd

  dat.vol <- xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Volume"))
  dat.iss <- xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//Journal//Issue"))
  dat.pag <- xmlToDataFrame(nodes=getNodeSet(pm_abstracts_xml,"//MedlinePgn"))

  #Vreemde correctie nodig, want 1 journal had "62 Suppl 2" als volume
  #dat.vol[["text"]] <- ifelse(grepl(".*Suppl.*", dat.vol[["text"]]), gsub("(\\d{2,3}).*", "\\1", dat.vol[["text"]]), dat.vol[["text"]])

  # dat.vol <- rbind(dat.vol, data.frame(text=rep(NA, length(unique(pm.table$jabbrv)))))
  # dat.iss <- rbind(dat.iss, data.frame(text=rep(NA, length(unique(pm.table$jabbrv)))))
  # dat.pag <- rbind(dat.pag, data.frame(text=rep(NA, length(unique(pm.table$jabbrv)))))


  pm.table$volume <- "EOP"
  pm.table$issue  <- "EOP"
  pm.table$pages  <- "EOP"

  pm.table$volume[wh.sel.abs] <- unname(unlist(apply(data.frame(dat.vol[["text"]], rle(pm.table$pmid[wh.sel.abs])$lengths), 1, FUN=function(x) rep(x[1], each=x[2]))))
  pm.table$issue[wh.sel.abs]  <- unname(unlist(apply(data.frame(dat.iss[["text"]], rle(pm.table$pmid[wh.sel.abs])$lengths), 1, FUN=function(x) rep(x[1], each=x[2]))))
  pm.table$pages[wh.sel.abs]  <- unname(unlist(apply(data.frame(dat.pag[["text"]], rle(pm.table$pmid[wh.sel.abs])$lengths), 1, FUN=function(x) rep(x[1], each=x[2]))))

  pm.table$journal <- NA
  pm.table$journal <- ifelse(pm.table$jabbrv %in% "Br J Clin Pharmacol", "British Journal of Clinical Pharmacology", pm.table$journal)
  pm.table$journal <- ifelse(pm.table$jabbrv %in% "Clin Pharmacokinet",  "Clinical Pharmacokinetics",                pm.table$journal)
  pm.table$journal <- ifelse(pm.table$jabbrv %in% "J Clin Pharmacol",    "Journal of Clinical Pharmacology",         pm.table$journal)
  pm.table$journal <- ifelse(pm.table$jabbrv %in% "Clin Pharmacol Ther", "Clinical Pharmacology & Therapeutics",     pm.table$journal)


  ## Fetch author names
  pm.table$firstname[grepl(".*\\s[A-z]$", pm.table$firstname)] <- paste0(pm.table$firstname[grepl(".*\\s[A-z]$", pm.table$firstname)], ".")
  pm.table$auth <- paste0(pm.table$firstname, " ",pm.table$lastname)

  pm.table$author <- NA

  for(set in unique(pm.table$pmid)){

    pm.table$author[pm.table$pmid %in% set] <- rep(paste0(pm.table$auth[pm.table$pmid %in% set], collapse = ", "), length=sum(pm.table$pmid %in% set))

  }


  ## Remove duplicated rows based on PMID, duplicated rows WAS necessary for creating one string of all authors.
  pm.table <- pm.table[!duplicated(pm.table$pmid),]


  ## create DOI
  pm.table$doi <- paste0("https://doi.org/", pm.table$doi)


  ## Set months
  pm.table <- merge(pm.table, pm.extract[,c("pmid","mon")], by="pmid", sort=F)
  pm.table$month <- match(pm.table$mon, ds.months$rep) #XML geeft andere data dan de abstract fetch bovenin, daarom is dit nodig!
  pm.table$mon.full <- month.name[as.numeric(pm.table$month)]
  pm.table$mon.abbr <- month.abb[as.numeric(pm.table$month)]


  ## Remove all HTML elements
  for(abstr in 1:length(pm.table$abstract)){
    pm.table$abstract[abstr] <- HTMLdecode(pm.table$abstract[abstr])
  }


  ## Remove points at the end of title
  pm.table$title[which(grepl("^.*\\.$", pm.table$title))] <- gsub("^(.*)\\.$", "\\1", pm.table$title[which(grepl("^.*\\.$", pm.table$title))])


  ## Remove all HTML elements from titles
  pm.table$title <- HTMLdecode(pm.table$title) # niet voldoende om alle elementen te verwijderen!

  remov.wrd <- c("<sub>", "</sub>", "<sup>", "</sup>")
  pm.table$title <- str_remove_all(pm.table$title, paste(remov.wrd, collapse = "|"))


  ## Remove all HTML elements from authors
  pm.table$author <- HTMLdecode(pm.table$author)


  ## Create table using regex of all data necessary for abstract which contain paragraphs (Introduction, methods, etc.)
  abstr.txt <- data.frame(text=as.character(), subj=as.character(), pmid=as.numeric())
  remov.wrd <- c("<sub>", "</sub>", "<sup>", "</sup>")


  for(i in 1:length(pm_entrez_id$IdList)){

    txt <- trimws(sapply(abstr.extract, custom_grep, tag = "Abstract", format = "char", USE.NAMES = FALSE)[i,])

    ch.jocp <- T %in% grepl("(ELocationID)|(MedlinePgn)|(ISOAbbreviation)|(PubmedData)", txt) #random keuze

    if(!ch.jocp){
      txt <- strsplit(txt, "><")[[1]]
    }
    #txt <- strsplit(txt, "\\s{3,}")[[1]]
    #txt <- strsplit(txt, "<Copy")

    #Journal of Clinical Pharmacology - check
    #gsub("^<AbstractText>(.*)</AbstractText>$", "\\1", txt[31])


    if(ch.jocp){
      txt <- strsplit(txt, "><")[[1]]
      txt <- if(T%in%grepl("<Abstract>", txt)){
        gsub('^.*<Abstract>(.*)', "\\1", txt)
      }else{
        txt[grepl("AbstractText Label.*", txt)|grepl("AbstractText>.*", txt)]
      }
    }

    abstr.par <- !grepl("<?AbstractText>.*</AbstractText>?", txt[1]) | #  Does abstract have paragraphs? Vraagteken is belangrijk als er toch een copyright paragraaf is
      T%in%grepl("NlmCategory", txt)  # in 1 van de paragrafen komt er een label of  Nlm categorie type XLM data voor?

    if(!ch.jocp){
      txt <- txt[!grepl("(<Copy)?rightInformation>", txt)] # Remove copyright rows
    }

    if(ch.jocp){
      txt <- gsub("^(.*)<CopyrightInformation>.*$", "\\1", txt)
    }

    if(!ch.jocp){
      txt <- txt[!grepl("CLINICAL TRIAL REGISTRY NUMBERS", txt)] # Remove copyright rows
    }

    abstr.txt.fill <- data.frame(text=as.character(), subj=as.character())

    if(abstr.par){

      for(j in 1:length(txt)){

        incl.text <- gsub('^<?AbstractText( Label=\\".*\\"(\\s)?(.*NlmCategory=\\".*\\")?)?>(.*)</AbstractText>?$', "\\4", txt[j])
        incl.subj <- gsub('.*Label=\\"(.*)\\"(\\s)?(.*NlmCategory.*)?.*', "\\1", txt[j])
        incl.subj <- tolower(incl.subj)


        incl.text <- HTMLdecode(incl.text)
        incl.subj <- paste(toupper(substr(incl.subj, 1, 1)), substr(incl.subj, 2, nchar(incl.subj)), sep="")

        incl.text <- str_remove_all(incl.text, paste(remov.wrd, collapse = "|"))

        abstr.txt.fill[j,] <- list(text=incl.text, subj=incl.subj)


      }

    }else{

      incl.text <- gsub('^<?AbstractText>(.*)</AbstractText>?$', "\\1", txt)
      incl.subj <- "Abstract"

      incl.text <- HTMLdecode(incl.text)

      incl.text <- str_remove_all(incl.text, paste(remov.wrd, collapse = "|"))

      abstr.txt.fill <- data.frame(text=incl.text, subj=incl.subj)
    }

    abstr.txt.fill$pmid <- as.numeric(pm_entrez_id$IdList[[i]])
    abstr.txt <- rbind(abstr.txt, abstr.txt.fill)
  }


  abstr.txt$seq <- cumsum(ave(1:nrow(abstr.txt), abstr.txt$pmid, FUN=function(x) {1:length(x)})<2)
  # Aparte abstract:
  # BJCP
  # CP

  #Calctlate issue of 4CP
  CP.issue <- data.frame(iss=seq(23,100,1))

  CP.issue$mon.num <- c(c(seq_along(CP.issue$iss)+5) %% 12) + 1
  CP.issue$mon.nam <- month.name[CP.issue$mon.num]

  CP.issue$year <- 2020+cumsum(CP.issue$mon.num==1)

  assign("CP.issue",   CP.issue,   envir = globalenv())
  assign("abstr.txt", abstr.txt , envir = globalenv())
  assign("pm.table",   pm.table,   envir = globalenv())
  assign("pm.extract", pm.extract, envir = globalenv())

}
