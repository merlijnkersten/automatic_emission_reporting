# This code requires the 'ta' data table.
# I do not understand what T_results_v02+ code is required to create it.
# A significant part of the code below (except for "Functions" and further) is unnecessary.
# Feel free to remove the unnecessary code as you see fit.

#---- Packages ----
# Which of these are required?
library(plyr)
library(tidyverse)
library(openxlsx)
require(googledrive)
library(gt)
#devtools::install_github("ianmoran11/mmtable2")
library(mmtable2)
library(dplyr)
library(stringr)
library(tidyr)
library(sjlabelled)
library(patchwork)
library(RColorBrewer)
library(paletteer)
library(officer)
library("colorspace")

# ---- Load t ----
# How much of this code is necessary?

options(scipen=100,digits=4)

if (.Platform$OS.type == 'windows') {
  Sys.setlocale(category = 'LC_ALL','English_United States.1250')
} else {
  Sys.setlocale(category = 'LC_ALL','en_US.UTF-8')
}

sx <- 7  # for EUA SA

#load("C:/Users/Merlijn Kersten/Documents/Univerzita Karlova/Before_Batch_Import.RData")
load("G:/Shared drives/MAKRO MODELY/TIMES/TIMES_V02+/Before_Batch_Import.RData")
path_fig <- c("C:/Users/czpkersten/Desktop")
path_tab <- c("C:/Users/czpkersten/Desktop")

##:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::batch import:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::-----
#setwd("C:/VEDA/GAMS_WrkTIMES_v02+/BEV_cars")
#setwd("C:/VEDA/GAMS_WrkTIMES_v02+/RegSim_preliminary")
setwd("C:/VEDA/GAMS_WrkTIMES")

imdir<-list.files(recursive = TRUE,pattern = ".vd")
imdvd<- imdir[str_detect(imdir,".vd$")]   # vd files
imdvde<-imdir[str_detect(imdir,".vde$")]   # vde files

#todays scenarios:

imdvd <- imdvd[str_detect(imdvd,format(Sys.Date(),"%d%m"))]
imdvde <- imdvde[str_detect(imdvde,format(Sys.Date(),"%d%m"))]


#yesterdays scenarios:

#imdvd <- imdvd[str_detect(imdvd,format(Sys.Date()-1,"%d%m"))]
#imdvde <- imdvde[str_detect(imdvde,format(Sys.Date()-1,"%d%m"))]


imdvd
imdvde

t <- NULL
for (n in 1:length(imdvd)) {
  #b<-unlist(strsplit(imdvd[n],c("/"))) # ("/","[.]")))
  a <- imdvd[n]%>%str_split("[.]") %>% unlist()
  tt<-read.csv(imdvd[n],skip = 13,header = F)
  names(tt)<-c("attribute","commodity","process","period","region","vintage","timeSlice","userConstraint","pv")
  tt$scenario<-a[1] %>%str_sub(1,-6)
  tt$attribute <- tt$attribute %>% as.factor()
  t<-rbind(t,tt)
}

d <- NULL
for (n in 1:length(imdvde)) {
  #b<-unlist(strsplit(imdvde[n],c("/"))) # ("/","[.]")))
  a <- imdvde[n] %>%str_split("[.]") %>% unlist()
  dd<-read.csv(imdvde[n],header = F)
  names(dd)<-c("type","region","code","description")
  dd$scenario<-a[1] %>%str_sub(1,-6)
  dd$type <- dd$type %>% as.factor()
  d<-rbind(d,dd)
}

d <- d %>% select(-region,-scenario)%>% unique()
p <- d %>% filter(type=="Process") %>% select(code,description)%>% dplyr::rename("process"="code","proc_desc"="description")
c <- d %>% filter(type=="Commodity")%>% select(code,description)%>% dplyr::rename("commodity"="code","com_desc"="description")
c$dup <- duplicated(c$commodity)
cdup <- c %>% filter(dup=="TRUE")
c <- c %>% filter(dup=="FALSE") %>% select(-dup)


t <- left_join(t,p)
t <- left_join(t,c)

#t <- rbind(t,select(nkep,names(t)))


t <- left_join(t,com[,c("commodity","unit","kg_m3","red.ii..kg_m3","red.ii..mj_kg","red.ii..mj_dm3")])
# all fuel to PJ and kt
t$pj <- ifelse(t$unit=="PJ",t$pv,NA)
t$pj <- ifelse(t$unit=="kt",t$pv*t$red.ii..mj_kg/1000,t$pj)
t$pj <- ifelse(t$unit=="m3",t$pv*t$red.ii..mj_dm3/1000000,t$pj)

t$kt <- ifelse(t$unit=="kt",t$pv,NA)
t$kt <- ifelse(t$unit=="PJ",t$pv/t$red.ii..mj_kg*1000,t$kt)
t$red.ii..kg_m3 <- t$red.ii..kg_m3 %>% as.numeric()
t$kt <- ifelse(t$unit=="m3",t$pv*t$red.ii..kg_m3/1000000,t$kt)

##---- sets ----

## commodity sets
comod <- t %>% filter(commodity!="-") %>% select(commodity)%>% unique()

comod_prim <- comod %>% filter(.,str_detect(comod$commodity,"^BIO|^COA|^GASNAT|^OIL|^RENGEO|RENWIN|RENHYD|RENSOL|^NUC|^ELCHIG")) %>% unique() %>% unlist()%>% as.vector() %>% sort()

comod_sec <- comod %>% filter(.,str_detect(comod$commodity,"^SUP|^ELC|^IND|TRALPG|TRAGSL|TRAKER|TRADST|TRAHFO|TRAGAS|TRAGH2|^COM|^RSD|^AGR"),
                              !str_detect(comod$commodity,"AGR$|FM|ELCHIG|ELCLOW|ELCMED")) %>% unique()
comod_sec <- comod_sec["commodity"] %>% unique()%>% unlist() %>% as.vector() %>% sort()

comod_ind <- comod %>% filter(.,str_detect(comod$commodity,"^IND")) %>% unique()
comod_ind <- comod_ind["commodity"] %>% unique()%>% unlist() %>% as.vector()

comod_pol <- comod %>% filter(.,str_detect(comod$commodity,"^PM_|^NOx_|^SOx_")) %>% unique()
comod_pol <- comod_pol["commodity"] %>% unique()%>% unlist() %>% as.vector()

comod_co2 <- comod %>% filter(.,str_detect(comod$commodity,"^C_")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
comod_ch4 <- comod %>% filter(.,str_detect(comod$commodity,"^CH4_")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
comod_n2o <- comod %>% filter(.,str_detect(comod$commodity,"^N2O_")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
comod_ghg <- c(comod_co2,comod_ch4,comod_n2o)


gr_comod_sec <- comod_sec %>% str_sub(.,start = 4) %>% unique()

elc <- comod %>% filter (.,str_detect(comod$commodity,"ELCHIG|ELCLOW|ELCMED|INDELC"))
elc <- elc["commodity"] %>% unique()%>% unlist() %>% as.vector()

comod_heat <- comod %>% filter (.,str_detect(comod$commodity,"HET")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
comod_indheat <- comod %>% filter (.,str_detect(comod$commodity,"^I"),str_detect(comod$commodity,"TH$")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
heat_out <- heat["commodity"] %>% unique()%>% unlist() %>% as.vector()
heat <- comod %>% filter (.,str_detect(comod$commodity,"HTH$|LTH$"))%>% unique()%>% unlist() %>% as.vector()

comod_ren_tra <- c("FAME1G","FAMEixB","FAME2G","FAEE1G","FAEEixB","MTBE2G","ETBE1G","ETOH1G","ETOH2G","HVO1G","HVOixB",
                   "HVO2G","BBEN1G","BBENixB","BLPG1G","BLPGixB","BCNG2G","H22G","BMN1G","BMN2G","TRAELCR","TRAELCT")

comod_ren <- comod %>% filter(.,str_detect(comod$commodity,"^BIO|^RENGEO|RENWIN|RENHYD|RENSOL")) %>% unique() %>% unlist()%>% as.vector() %>% sort()
comod_h2 <- comod %>% filter (.,str_detect(comod$commodity,"H2"),!str_detect(comod$commodity,"^s"),commodity!="ETOH2G")%>% unique() %>% unlist()%>% as.vector() %>% sort()


comod_heat_chp <- comod %>% filter(.,str_detect(comod$commodity,"^HET"),!str_detect(comod$commodity,"TH$")) %>% unique() %>% unlist()%>% as.vector() %>% sort()

comod_ets_by <- comod %>%  filter(.,str_detect(comod$commodity,"^ets")) %>% unique()%>% unlist() %>% as.vector()



## process sets
proc_dummy <- t %>% filter(.,str_detect(t$process,"IMPDEMZ|IMPMATZ|IMPNRGZ"))
proc_dummy  <- proc_dummy ["process"] %>% unique()%>% unlist() %>% as.vector()

proc_prim <- t %>% filter(.,str_detect(t$process,"^IMP|^EXP|^MIN|^SPI|^SPR"))
gr_proc_prim <- proc_prim$process %>% str_sub(.,start = 1,end = 3) %>% unique()
proc_prim <- proc_prim["process"] %>% unique()%>% unlist() %>% as.vector()

proc_trans <- t %>% filter(.,str_detect(t$process,"TRANS"))
proc_trans  <- proc_trans ["process"] %>% unique()%>% unlist() %>% as.vector()

proc_gfec <- t %>% filter(attribute=="VAR_FIn",commodity %in% c(comod_prim,elc,"HETHTH","HETLTH","BAF","NMF"), str_detect(t$process,"00$|01$"),
                          !str_detect(t$process,"^NEC|^NEO|^SPIP|^SPRI|^SRET|^SSC|^STRAN|^ELC|^EU"))
proc_gfec <- proc_gfec["process"]%>% unique()%>% unlist() %>% as.vector()

proc_elc <- t %>% filter(commodity %in% elc, attribute=="VAR_FOut", process %!in% proc_trans,!str_detect(t$process,"^CC|BTRY|STG|CAES|HYDP|INDELC|IMP"))
proc_elc <- proc_elc["process"]%>% unique()%>% unlist() %>% as.vector()

proc_heat_sec <- t %>% filter(attribute=="VAR_FIn",commodity %in% c("HETHTH","HETLTH"), str_detect(t$process,"00$|01$"))
proc_heat_sec <- proc_heat_sec["process"]%>% unique()%>% unlist() %>% as.vector()

proc_gas_dist <- t %>% filter(attribute=="VAR_FIn",commodity=="GASNAT", process!="NECDEMAND00" ) %>% select(process)%>% unique() %>% unlist() %>% as.vector()

proc_res_e <- t %>% filter(.,str_detect(t$process,"EUHYD|PVSOL|EUWIN|EUGEO|PUGEO|EUFSC|EAUTHYD|EAUTSOL|EAUTWIN|EAUTGEO"),
                           !str_detect(t$process,"^CC|BTRY|STG|CAES|HYDPS|HYDPUMP"))
proc_res_e <- proc_res_e["process"]%>% unique()%>% unlist() %>% as.vector()

proc_bio_elc <- ta %>% filter(str_detect(ta$process,"BGS|BIO|WOO"),str_detect(ta$commodity,"ELCLOW|ELCMED|ELCHIG|INDELC|COMELC"),attribute=="VAR_FOut")
proc_bio_elc <- proc_bio_elc["process"]%>% unique()%>% unlist() %>% as.vector()

proc_mun_elc <- ta %>% filter(str_detect(ta$process,"MUN|WST"),str_detect(ta$commodity,"ELCLOW|ELCMED|ELCHIG|INDELC|COMELC"),attribute=="VAR_FOut")
proc_mun_elc <- proc_mun_elc["process"]%>% unique()%>% unlist() %>% as.vector()

proc_bio_heat <- ta %>% filter(str_detect(ta$process,"BGS|BIO|WOO"),str_detect(ta$commodity,"HTH|LTH"),attribute=="VAR_FOut")%>% unique()%>% unlist() %>% as.vector()


proc_indbio <- ta %>% filter(commodity=="INDBIO",attribute=="VAR_FIn")
proc_indbio <- proc_indbio["process"]%>% unique()%>% unlist() %>% as.vector()


proc_crf_com <- ta %>% filter(str_detect(ta$commodity,"^C_"),attribute=="VAR_FOut") %>% select(process,commodity) %>% unique() %>%
  filter(commodity!="C_tech", !(process=="ICH6"& commodity=="C_1A4ai"))

proc_crf_com$crf <- proc_crf_com$commodity %>% str_sub(3,8)
proc_crf_com$crf_short <- case_when(str_detect(proc_crf_com$crf,"1A1")~ "1A1",
                                    str_detect(proc_crf_com$crf,"1A4")~ "1A4",
                                    str_detect(proc_crf_com$crf,"1B")~ "1B",
                                    str_detect(proc_crf_com$crf,"^2")~ "2",
                                    TRUE~ proc_crf_com$crf)
proc_crf_com <- proc_crf_com %>% select(-commodity)
proc_crf <- proc_crf_com ["process"]%>% unique()%>% unlist() %>% as.vector()

proc_cghg <- ta %>% filter(commodity=="CGHG",attribute=="VAR_FOut") %>% select(process) %>% unique()%>% unlist() %>% as.vector()

proc_Troad <- c(proc_TBUS,proc_TCARS,proc_TCOACH,proc_TLUV,proc_TMOT,proc_TNA)

proc_imp_exp <- c(proc_IMPORT_PR,proc_EXPORT_PR)

comod_prim
comod_sec

proc_prim

proc_h2 <- ta %>% filter(commodity %in% comod_h2,attribute=="VAR_FOut", !process %>% str_detect('^ST|P_ST')) %>% select(process) %>% unique()%>% unlist() %>% as.vector()

gr_comod_sec
gr_proc_prim


att_cost <- t$attribute %>% unique() %>%  str_subset(.,"Cost")
att_varM <- t$attribute %>% unique() %>%  str_subset(.,"VAR_") %>%  str_subset(.,"M$")
att_var <- t$attribute %>% unique() %>%  str_subset(.,"VAR_")
att_var <- att_var[att_var %!in% att_varM]
att_io <- c('VAR_FIn','VAR_FOut')

t$attribute %>% as.character()%>% filter(!str_detect(t$attribute,"M$"),str_detect(t$attribute,"VAR_"))

# --- --- --- --- ---
# ---- Functions ----
# --- --- --- --- ---

# Use openxlsx for reading (loadWorkbook), writing (writeData) and saving (saveWorkbook) Excel files
# Documentation: https://cran.r-project.org/web/packages/openxlsx/openxlsx.pdf
library(openxlsx)


get_commodity_value <- function(i, column) {
  # Get the pollution values for every NFR code in i for the specified column.
  
  format_commodity <- function(j, column) {
    # Create a vector of formatted commodities from a vector of commodity codes and a column.

    vector <- c()
    
    for (x in j) {
      commodity <- paste(column, x, sep="_")
      vector <- append(vector, commodity)
    }
    
    return(vector)
    
  }
  
  i_formatted <- format_commodity(i, column)
  
  # Looks up the PV value for every commodity in 'i'
  
  find_value <- function(x) {
    # Get PV sum of VAR_FOut values in NKEP_high in 2025 for commodity x.
    # PARAMETER
    value <- (ta %>% filter(commodity==x, 
            scenario=="NKEP_high", 
            period==2025, 
            attribute=="VAR_FOut"))$pv %>% sum()
    return(value)
  }
  
  value_sum = 0
  
  for (y in i_formatted) {
    # Find value for every commodity x in vector i (also works if i is a single commodity code character)
    value_sum <- value_sum + find_value(y)
  }
  
  return(value_sum)
  
}


alpha_num <- function(i) {
  # Translate an Excel column string (e.g. AC) into a column number (AC => 26 + 3 = 29)
  
  a_to_1 <- function(a) {
    # Get the number of a single letter of the alphabet
    alphabet <- "abcdefghijklmnopqrstuvwxyz"
    #             0000000011111111112222222
    #            12345678901234567890123456
    # Find R version of Python's 'find()'
    lst <- unlist(strsplit(alphabet, split=""))
    #lst <- list("a", "b", "c", ..., "z")
    number <- match(tolower(a), lst)
    return(number)
  }
  
  
  if (nchar(i) == 1) {
    
    return(a_to_1(i))
  
  } else if (nchar(i)==2) {
    
    i_1 <- substr(i, 1, 1)
    i_2 <- substr(i, 2, 2)
    return(a_to_1(i_1)*26 + a_to_1(i_2)) 
    # Python: a_to_1(i[0])*26 + a_to_1(i[1]) --- how to do this in R (not object-oriented)?
    
  } else {
  
    len = nchar(i)
    stop(paste("Length of ", len, " not supported yet ('", i, "')", sep=""))
  
  }
}


# List of commodities in VEDA
commodities <- ta$commodity %>% unique() %>% sort()

# List of NFR codes in Excel files (note: codes do not exist for all pollutants)
# PARAMETER
rows_in_veda <- c("1A1a", "1A1b", "1A1c", "1A2", "1A3", "1A4ai", "1A4bi", "1A4ci", 
                  "1A4cii", "1B1", "1B2", "2A1", "2A2", "2A3", "2A4", "2C1", "2D1")

# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
# -------- File A: annex_iV_rev2022_v1 ------------------------------
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
# The code for this file has many comments explaining its method.

# Load Excel file
# PARAMETER
path_excel ="C:/Users/czpkersten/Documents/automatic_emission_reporting/annex_iV_rev2022_v1.xlsx"
wb <- loadWorkbook(path_excel)

# Pollutants and their respective Excel columns
# PARAMETER
columns_a <- list(
  # pollutant (VEDA name) = column in excel file
  "NOx" = "E",
  "SOx" = "G",
  "PM"  = "K") # TSP

# Load rows of file A (from a separate CSV)
# PARAMETER
rows_a <- read.csv("C:/Users/czpkersten/Documents/automatic_emission_reporting/rows_a.csv", header=FALSE)$V1

# These rows don't exist in VEDA but a similar code does exist
# PARAMETER
edge_rows_a <- list(#code in Excel file = code in VEDA
                    "1A2a"     = "1A2",
                    "1A3ai(i)" = "1A3",
                    "1B1a"     = "1B1")

for (column in names(columns_a)) {
  # For every column in the Excel file
  values <- c()
  
  for (row in rows_a) {
    # And for every row in the Excel file
    
    if (row %in% names(edge_rows_a)) {
      row <- edge_rows_a[row]
    }
    
    # Create full commodity name (as in VEDA): <pollutant>_<NFR code> (e.g, C_1A1a)
    commodity <- paste(column, row, sep="_")
    
    if (commodity %in% commodities) {
      # If the commodity exists in the VEDA file, find its value
      commodity_value <- get_commodity_value(row, column)
    
    } else {
      # Else, return NA
      commodity_value <- NA
    
    }
    
    # Save every row value to the values vector
    values <- append(values, commodity_value)
  }
  
  # Find the number of the column (e.g. AC = 26 + 3 = 29) and save the values list to the correct position in the Excel file
  col <- alpha_num(columns_a[column])
  # PARAMETER
  writeData(wb, "2025_WM", values, startCol=col, startRow=14, colNames=FALSE, rowNames=FALSE  )

  }

# Save the Excel file
# PARAMETER
path = "C:/Users/czpkersten/Documents/automatic_emission_reporting/File a output.xlsx"
saveWorkbook(wb, path, overwrite=TRUE)


# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- ---
# -------- File B: GovReg_Proj_T1a_T1b_T5a_T5b_v1.1 -----------------
# --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- --- 
# Broadly similar to File A; I only added comments where they differ.

# PARAMETER
path_excel ="C:/Users/czpkersten/Documents/automatic_emission_reporting/GovReg_Proj_T1a_T1b_T5a_T5b_v1.1 - Table1a.xlsx"
wb <- loadWorkbook(path_excel)

# PARAMETER
columns_b <- list(
  # pollutant = column in excel file (2025)
  "C"   = "K",
  "CH4" = "AJ", 
  "N20" = "BI"
)

# PARAMETER
rows_b <- read.csv("C:/Users/czpkersten/Documents/automatic_emission_reporting/rows_b.csv", header=FALSE)$V1

# PARAMETER
edge_rows_b = list(#code in Excel file = code in VEDA
                   "1A4a" = "1A4ai",
                   "1A4b" = "1A4bi")

for (column in names(columns_b)) {
  values <- c()
  for (row in rows_b) {
    
    if (row %in% names(edge_rows_b)) {
      row <- edge_rows_b[row]
    }
    
    commodity <- paste(column, row, sep="_")
    
    if (row=="1") {
      # File B has a few "summary" columns (such as 1), filling these requires summing the values for every
      # commodity in rows_vector.
      # PARAMETER
      rows_vector <- c("1A1a", "1A1b", "1A1c", "1A2", "1A3", "1A4ai", "1A4bi", "1A4ci", "1A4cii", "1B1", "1B2")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="1A") {
      # PARAMETER
      rows_vector = c("1A1a", "1A1b", "1A1c", "1A2", "1A3", "1A4ai", "1A4bi", "1A4ci",  "1A4cii")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="1A1") {
      # PARAMETER
      rows_vector <- c("1A1a", "1A1b", "1A1c")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="1A4") {
      # PARAMETER
      rows_vector <- c("1A4ai", "1A4bi", "1A4ci", "1A4cii")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="1A4c") {
      # PARAMETER
      rows_vector <- c("1A4ci", "1A4cii")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="1B") {
      # PARAMETER
      rows_vector <- c("1B1", "1B2")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="2") {
      # PARAMETER
      rows_vector <- c("2A1", "2A2", "2A3", "2A4", "2C1", "2D2")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (row=="2A") {
      # PARAMETER
      rows_vector <- c("2A1", "2A2", "2A3", "2A4")
      commodity_value <- get_commodity_value(rows_vector, column)
      
    } else if (commodity %in% commodities) {
      
      commodity_value <- get_commodity_value(row, column)
      
    } else {
      
      commodity_value <- NA
    
    }
    
    values <- append(values, commodity_value)
    
  }

  col <- alpha_num(columns_b[column])
  # PARAMETER
  writeData(wb, "Table1a", values, startCol=col, startRow=19, colNames=FALSE, rowNames=FALSE)

}

# PARAMETER
path = "C:/Users/czpkersten/Documents/automatic_emission_reporting/File b output.xlsx"
saveWorkbook(wb, path, overwrite=TRUE)

