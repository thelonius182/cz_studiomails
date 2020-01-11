library(googledrive)
library(keyring)
library(readxl)
library(readr)
library(yaml)
library(stringr)
library(dplyr)
library(lubridate)
library(gmailr)

# define functions ----
cz_extract_sheet <- function(ss_name, sheet_name) {
  read_xlsx(ss_name,
            sheet = sheet_name,
            .name_repair = ~ ifelse(nzchar(.x), .x, LETTERS[seq_along(.x)]))
}

cz_get_url <- function(cz_ss) {
  cz_url <- paste0("url_", cz_ss)
  
  # use [[ instead of $, because it is a variable, not a constant
  paste0("https://", config$url_pfx, config[[cz_url]]) 
}

# init config ----
config <- read_yaml("config.yaml")

# downloads GD ----

# aanmelden bij GD loopt via de procedure die beschreven is in "Operation Missa > Voortgang > 
# 4.Toegangsrechten GD".

# Roosters 3.0 ophalen bij GD
path_roosters <- paste0(config$gs_downloads, "/", "roosters.xlsx")
drive_download(file = cz_get_url("roosters"), overwrite = T, path = path_roosters)

# sheets als df ----

tbl_raw_presentatie <- cz_extract_sheet(path_roosters, sheet_name = "presentatie")

# refactor raw tables 
tbl_presentatie.I <-  tbl_raw_presentatie %>% 
  filter(row_number() > 1) %>% 
  select(uitzending = B, uren, pres = Presentatie, tech = Techniek, live = Status) %>% 
  mutate(live = if_else(live == "Live", T, F))

# teamleden: via SalesForce
path_teamleden <- paste0(config$gs_downloads, "/studioteamleden.csv")
tbl_teamleden <- read_csv(file = path_teamleden,
                            col_types = cols(NPE01__ALTERNATEEMAIL__C = col_skip())
  ) %>% 
  mutate(tlf = if_else(!is.na(MOBILEPHONE), MOBILEPHONE, PHONE)) %>% 
  select(-MOBILEPHONE, -PHONE)

tbl_presentatie <- tbl_presentatie.I %>% 
  left_join(tbl_teamleden, by = c("pres" = "NAME")) %>% 
  left_join(tbl_teamleden, by = c("tech" = "NAME")) %>% 
  filter(live & (TEAMSERVICE_ONTVANGERS__C.x | TEAMSERVICE_ONTVANGERS__C.y))

rm(tbl_presentatie.I, tbl_raw_presentatie, tbl_teamleden)

# overmorgen
overmorgen <- today(tzone = "Europe/Amsterdam") + ddays(2L)
mailen_naar_t2 <- tbl_presentatie %>% 
  filter(uitzending == overmorgen)

# over 10 dagen
over_10_dagen <- today(tzone = "Europe/Amsterdam") + ddays(10L)
mailen_naar_t10 <- tbl_presentatie %>% 
  filter(uitzending == over_10_dagen)

