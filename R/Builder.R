
#' Build the actual word file for publication. This word file still needs to be compiled to a PDF before sending.
#'
#' @param wd.data Path to data
#'
#' @return
#' @export
#'
#' @examples
#'
#'

Builder <- function(wd.data){

  library(officer)
  library(magrittr)
  library(textutils)

  #rm(list=ls()[!(ls() %in% c("CP.issue","abstr.txt","pm.table","pm.extract"))])

  if(!exists("wd.data")){wd.data <- "C:/Users/TPreijers/Dropbox/4Abstracts/R_test"}

  CP.issue   <- get0("CP.issue",   ifnotfound = "CP.issue not found",   envir = .GlobalEnv)
  abstr.txt  <- get0("abstr.txt",  ifnotfound = "abstr.txt not found",  envir = .GlobalEnv)
  pm.extract <- get0("pm.extract", ifnotfound = "pm.extract not found", envir = .GlobalEnv)
  pm.table   <- get0("pm.table",   ifnotfound = "pm.table not found",   envir = .GlobalEnv)

  wh.ord     <- order(pm.table$journal)
  pm.table   <- pm.table[wh.ord,]
  pm.extract <- pm.extract[wh.ord,]

  pm.lit     <- pm.table
  pm.lit$seq <- ave(1:nrow(pm.lit), pm.lit$jou, FUN=function(x) 1:length(x))

  pm.jou     <- pm.table[!duplicated(pm.table$jou),]

  pm.jou$issue  <- pm.extract$iss[!duplicated(pm.extract$jou)]
  pm.jou$volume <- pm.extract$vol[!duplicated(pm.extract$jou)]
  pm.jou$pages  <- pm.extract$pag[!duplicated(pm.extract$jou)]


  ##overschrijf PM.table
  pm.table$pages  <- pm.extract$pag
  pm.table$issue  <- pm.extract$iss
  pm.table$volume <- pm.extract$vol


  text.norm <- fp_text(
    color = "black",
    font.size = 10,
    bold = FALSE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )
  hex_col_rose <- "#F90C8E"
  text.norm2 <- fp_text(
    color = hex_col_rose,
    font.size = 10,
    bold = TRUE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )
  text.norm3 <- fp_text(
    color = hex_col_rose,
    font.size = 11,
    bold = TRUE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )

  text.norm4 <- fp_text(
    color = "black",
    font.size = 12,
    bold = TRUE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )
  text.norm5 <- fp_text(
    color = "#004850",
    font.size = 10,
    bold = FALSE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )
  text.norm6 <- fp_text(
    color = "black",
    font.size = 10,
    bold = TRUE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )

  hyperlink <- fp_text(
    color = "#004850",
    font.size = 10,
    bold = FALSE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )

  hyperlink_BACK <- fp_text(
    color = "#004850",
    font.size = 10,
    bold = TRUE,
    italic = FALSE,
    underlined = FALSE,
    font.family = "Arial",
    vertical.align = "baseline",
    shading.color = "transparent"
  )

  ## Load docs!
  my_doc  <- read_docx(path = file.path(wd.data, "Template/4CP - empty.docx"))
  my_back <- read_docx(path = file.path(wd.data, "Template/4CP - empty2.docx"))
  my_intl <- read_docx(path = file.path(wd.data, "Internal_links/4CP - internal_links.docx"))

  my_back <- my_back %>% cursor_begin() %>% body_remove()
  my_intl <- my_intl %>% cursor_begin() %>% body_remove()

  #styles_info(my_doc)

  #20210319 tijdelijk voor deze issue, kan hierna weg
  # pm.lit[pm.lit$pmid=="33622228",]$journal <- "Clinical Pharmacokinetics"
  # pm.lit$seq[16] <- 4
  # pm.jou[pm.jou$pmid=="33622228",]$journal <- "Clinical Pharmacokinetics"



  ## Replace and create issue line
  txt.issue <- paste0("Issue ",
                      CP.issue$iss[CP.issue$mon.num %in% as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%m")) &
                                     CP.issue$year %in% as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%Y"))],
                      ", ",
                      month.name[as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%m"))], " ",
                      as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%Y")))


  my_doc <- my_doc %>% cursor_reach(keyword = "REPLACE") %>% ## Replace and write issue line
    body_add_fpar(fpar(ftext(txt.issue, text.norm2), fp_t = text.norm2),  pos = "on")



  ### Create first-page lines
  for(i in 1:nrow(pm.jou)){

    if(pm.jou$journal[i] %in% "British Journal of Clinical Pharmacology"){
      my_doc <- my_doc %>% cursor_reach(keyword = "British Journal of Clinical Pharmacology") %>%
        body_add_fpar(fpar(ftext(paste0(pm.jou$journal[i],", ", pm.jou$mon.full[i]," ",pm.jou$year[i],", ","Volume ",pm.jou$volume[i],", ","Issue ",pm.jou$issue[i]),
                                 text.norm),
                           fp_p = fp_par(padding.top=6, text.align = "justify"), fp_t = text.norm),  pos = "on")
    }

    if(pm.jou$journal[i] %in% "Clinical Pharmacokinetics"){
      my_doc <- my_doc %>% cursor_reach(keyword = "Clinical Pharmacokinetics") %>%
        body_add_fpar(fpar(ftext(paste0(pm.jou$journal[i],", ", pm.jou$mon.full[i]," ",pm.jou$year[i],", ","Volume ", pm.jou$volume[i],", ","Issue ",pm.jou$issue[i]),
                                 text.norm),
                           fp_p = fp_par(padding.top=6, text.align = "justify"), fp_t = text.norm),  pos = "on")
    }

    if(pm.jou$journal[i] %in% "Journal of Clinical Pharmacology"){
      my_doc <- my_doc %>% cursor_reach(keyword = "Journal of Clinical Pharmacology2") %>%
        body_add_fpar(fpar(ftext(paste0(pm.jou$journal[i],", ", pm.jou$mon.full[i]," ",pm.jou$year[i],", ","Volume ",pm.jou$volume[i],", ","Issue ",pm.jou$issue[i]),
                                 text.norm),
                           fp_p = fp_par(padding.top=6, text.align = "justify"), fp_t = text.norm),  pos = "on")
    }

    if(pm.jou$journal[i] %in% "Clinical Pharmacology & Therapeutics"){
      my_doc <- my_doc %>% cursor_reach(keyword = "Clinical Pharmacology & Therapeutics") %>%
        body_add_fpar(fpar(ftext(paste0(pm.jou$journal[i],", ", pm.jou$mon.full[i]," ",pm.jou$year[i],", ","Volume ",pm.jou$volume[i],", ","Issue ",pm.jou$issue[i]),
                                 text.norm),
                           fp_p = fp_par(padding.top=6, text.align = "justify"), fp_t = text.norm),  pos = "on")
    }
  }



  ## Content-page
  my_doc <- my_doc %>% cursor_end() %>%
    body_add_par("Content", style = "heading 2", pos="on") %>% body_bookmark("Content") ## _Content_1 is linked to BACK-button on each page


  for(i in 1:nrow(pm.lit)){

    # my_intl2 <- read_docx(path = "/Users/TPreijers/Dropbox/4Abstracts/R_test/Internal_links/4CP - internal_links.docx") %>% cursor_begin() %>% body_remove()
    #
    # for(nms in 1:20){
    #   if(nms==i){next}
    #   my_intl2 <- my_intl2 %>% cursor_reach(keyword = paste0("IL_", nms)) %>% body_remove() %>% cursor_begin()
    # }
    #
    # #docx_show_chunk(my_intl2)
    # my_intl2 <- my_intl2 %>% body_replace_all_text(paste0("IL_",i), pm.lit$title[i], perl=T)
    #
    # if(file.exists("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx")){
    #   file.remove("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx")
    # }
    # print(my_intl2, target = paste0("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx"))
    #

    ## Add empty line after each journal headings section
    if(pm.lit$seq[i]==1 & i>1){my_doc <- my_doc %>% cursor_end() %>% body_add_par(" ")}


    if(pm.lit$journal[i] %in% "British Journal of Clinical Pharmacology"){
      if(pm.lit$seq[i]==1){

        my_doc <- my_doc %>% cursor_end() %>%
          body_add_fpar(fpar(ftext(paste0(pm.lit$journal[i],", ", pm.lit$mon.full[i]," ", pm.lit$year[i],", ","Volume ", pm.lit$volume[i],", ","Issue ",pm.lit$issue[i]),
                                   text.norm3),
                             fp_p = fp_par(padding.top=10, text.align = "justify"), fp_t = text.norm3))
      }

      # my_doc <- my_doc %>% cursor_end() %>%
      #   body_add_docx(src = paste0("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx"))
      #
      my_doc <- my_doc %>% cursor_end() %>%
        body_add_fpar(fpar(ftext(paste0(pm.lit$title[i]), text.norm5), fp_p = fp_par(padding.top=6, padding.bottom=6, text.align = "justify"), fp_t = text.norm5), style="Bullet")
    }

    if(pm.lit$journal[i] %in% "Clinical Pharmacokinetics"){
      if(pm.lit$seq[i]==1){

        my_doc <- my_doc %>% cursor_end() %>%
          body_add_fpar(fpar(ftext(paste0(pm.lit$journal[i],", ", pm.lit$mon.full[i]," ",pm.lit$year[i],", ","Volume ",pm.lit$volume[i],", ","Issue ",pm.lit$issue[i]),
                                   text.norm3),
                             fp_p = fp_par(padding.top=10, text.align = "justify"), fp_t = text.norm3))
      }

      # my_doc <- my_doc %>% cursor_end() %>%
      #   body_add_docx(src = paste0("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx"))

      my_doc <- my_doc %>% cursor_end() %>%
        body_add_fpar(fpar(ftext(paste0(pm.lit$title[i]), text.norm5), fp_p = fp_par(padding.top=6, padding.bottom=6, text.align = "justify"), fp_t = text.norm5), style="Bullet")
    }

    if(pm.lit$journal[i] %in% "Clinical Pharmacology & Therapeutics"){
      if(pm.lit$seq[i]==1){

        my_doc <- my_doc %>% cursor_end() %>%
          body_add_fpar(fpar(ftext(paste0(pm.lit$journal[i],", ", pm.lit$mon.full[i]," ",pm.lit$year[i],", ","Volume ",pm.lit$volume[i],", ","Issue ",pm.lit$issue[i]),
                                   text.norm3),
                             fp_p = fp_par(padding.top=10, text.align = "justify"), fp_t = text.norm3))
      }

      # my_doc <- my_doc %>% cursor_end() %>%
      #   body_add_docx(src = paste0("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx"))

      my_doc <- my_doc %>% cursor_end() %>%
        body_add_fpar(fpar(ftext(paste0(pm.lit$title[i]), text.norm5), fp_p = fp_par(padding.top=6, padding.bottom=6, text.align = "justify"), fp_t = text.norm5), style="Bullet")

    }

    if(pm.lit$journal[i] %in% "Journal of Clinical Pharmacology"){
      if(pm.lit$seq[i]==1){

        my_doc <- my_doc %>% cursor_end() %>%
          body_add_fpar(fpar(ftext(paste0(pm.lit$journal[i],", ", pm.lit$mon.full[i]," ",pm.lit$year[i],", ","Volume ",pm.lit$volume[i],", ","Issue ",pm.lit$issue[i]),
                                   text.norm3),
                             fp_p = fp_par(padding.top=10, text.align = "justify"), fp_t = text.norm3))
      }

      # my_doc <- my_doc %>% cursor_end() %>%
      #   body_add_docx(src = paste0("/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP_temp_IL.docx"))

      my_doc <- my_doc %>% cursor_end() %>%
        body_add_fpar(fpar(ftext(paste0(pm.lit$title[i]), text.norm5), fp_p = fp_par(padding.top=6, padding.bottom=6, text.align = "justify"), fp_t = text.norm5), style="Bullet")
    }
  }


  ## Pagebrake on the end of the Contents-page
  my_doc <- my_doc %>% cursor_end() %>%
    body_add_break(pos = "after")



  ### Write abstract pages
  for(j in 1:nrow(pm.table)){


    my_doc <- my_doc %>% cursor_end() %>%
      body_add_fpar(fpar(ftext(pm.table$journal[j], text.norm4), fp_p = fp_par(padding.top=6, padding.bottom=6, text.align = "justify"), fp_t = text.norm4)) %>%
      body_add_par(pm.table$title[j], style="heading 2") %>%
      body_add_fpar(fpar(ftext(pm.table$author[j], text.norm), fp_p = fp_par(padding.top=0, text.align = "justify"), fp_t = text.norm)) %>%
      body_add_par("") %>%
      body_add_fpar(fpar(ftext(paste0(pm.table$mon.full[j]," ", pm.table$year[j],", ","Volume ", pm.table$volume[j],", ","Issue ",pm.table$issue[j],
                                      ", pages ", pm.table$pages[j]), text.norm), fp_p = fp_par(padding.top=0, text.align = "justify"), fp_t = text.norm)) %>%
      body_add_par("")


    if(T %in% is.na(abstr.txt$subj[abstr.txt$pmid %in% pm.table$pmid[j]])){

      my_doc <- my_doc %>% cursor_end() %>%
        body_add_fpar(fpar(ftext(abstr.txt$text[abstr.txt$pmid %in% pm.table$pmid[j]], text.norm), fp_p = fp_par(padding.top=0, text.align = "justify"), fp_t = text.norm)) %>%
        body_add_par("") %>%
        #body_add_par("") %>%
        body_add_fpar(fpar(hyperlink_ftext(text=pm.table$doi[j], href=pm.table$doi[j], prop=hyperlink))) %>%
        body_add_par("") %>%
        body_add_fpar(fpar(ftext("BACK", hyperlink_BACK), fp_p = fp_par(padding.top=0, text.align = "justify"), fp_t = hyperlink_BACK)) %>%
        #body_add_docx(src="/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP - empty2.docx") %>%
        body_add_break(pos = "after")
    }

    if(F %in% is.na(abstr.txt$subj[abstr.txt$pmid %in% pm.table$pmid[j]])){

      for(i in 1:length(abstr.txt$subj[abstr.txt$pmid %in% pm.table$pmid[j]])){

        my_doc <- my_doc %>% cursor_end() %>%
          body_add_par(abstr.txt$subj[abstr.txt$pmid %in% pm.table$pmid[j]][i], style="Bold") %>%
          body_add_par(abstr.txt$text[abstr.txt$pmid %in% pm.table$pmid[j]][i], style="Normal")

      }

      my_doc <- my_doc %>% cursor_end() %>%
        body_add_par("") %>%
        #body_add_par("") %>%
        body_add_fpar(fpar(hyperlink_ftext(text=pm.table$doi[j], href=pm.table$doi[j], prop=hyperlink))) %>%
        body_add_par("") %>%
        body_add_fpar(fpar(ftext("BACK", hyperlink_BACK), fp_p = fp_par(padding.top=0, text.align = "justify"), fp_t = hyperlink_BACK)) %>%
        #body_add_docx(src="/Users/TPreijers/Dropbox/4Abstracts/R_test/Template/4CP - empty2.docx") %>%
        body_add_break(pos = "after")
    }
  }


  my_doc <- my_doc %>% cursor_end() %>% body_remove()


  iss.mon <- CP.issue$mon.nam[CP.issue$mon.num %in% as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%m")) &
                                CP.issue$year %in% as.numeric(format(as.Date(Sys.Date(),format="%Y-%m-%d"),"%Y"))]


  ## Print document
  if(file.exists(file.path(wd.data, gsub("-", "", Sys.Date()),"_4CP_",iss.mon,".docx"))){
    file.remove(file.path(wd.data,  gsub("-", "", Sys.Date()),"_4CP_",iss.mon,".docx"))
  }

  print(my_doc, target = file.path(wd.data, paste0(gsub("-", "", Sys.Date()),"_4CP_",iss.mon,".docx")))
}
