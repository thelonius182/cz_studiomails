library(googledrive)
library(keyring)
library(readxl)
library(readr)
library(yaml)
library(stringr)
library(dplyr)
library(lubridate)

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
  filter(live & (TEAMSERVICE_ONTVANGERS__C.x | TEAMSERVICE_ONTVANGERS__C.y)) %>% 
  mutate(pres_roep = FIRSTNAME.x,
         pres_mail = EMAIL.x,
         pres_tlf = tlf.x,
         pres_send = TEAMSERVICE_ONTVANGERS__C.x,
         tech_roep = FIRSTNAME.y,
         tech_mail = EMAIL.y,
         tech_tlf = tlf.y,
         tech_send = TEAMSERVICE_ONTVANGERS__C.y
  )

rm(tbl_presentatie.I, tbl_raw_presentatie, tbl_teamleden)

source("src/send_mail.R", encoding = "UTF-8")

# overmorgen
overmorgen <- today(tzone = "Europe/Amsterdam") + ddays(2L)

mailen_dag2_pres <- tbl_presentatie %>% 
  select(pres_mail, pres_roep, uitzending, pres_send, uren, tech, tech_tlf) %>% 
  filter(uitzending == overmorgen & pres_send)

t1 <- reminder_text(pres = mailen_dag2_pres$pres_roep,
                    wanneer = "Overmorgen",
                    tech = mailen_dag2_pres$tech,
                    tlf = mailen_dag2_pres$tech_tlf,
                    hoe_laat = mailen_dag2_pres$uren,
                    uitz = format(mailen_dag2_pres$uitzending, "%A %d %B"))

send_reminder(pt_mail = "lcavdakker@gmail.com", reminder_txt = t1)

mailen_dag2_tech <- tbl_presentatie %>% 
  select(tech_mail, tech_roep, uitzending, uren, pres, pres_tlf) %>% 
  filter(uitzending == overmorgen & tech_send)

saveRDS(mailen_dag2_tech, file = paste0(config$cz_rds_store, "mailen_dag2_tech.RDS"))

# over 10 dagen
over_10_dagen <- today(tzone = "Europe/Amsterdam") + ddays(10L)

mailen_dag10_pres <- tbl_presentatie %>% 
  select(pres_mail, pres_roep, uitzending, uren, tech, tech_tlf) %>% 
  filter(uitzending == over_10_dagen & pres_send)

saveRDS(mailen_dag10_pres, file = paste0(config$cz_rds_store, "mailen_dag10_pres.RDS"))

mailen_dag10_tech <- tbl_presentatie %>% 
  select(tech_mail, tech_roep, uitzending, uren, pres, pres_tlf) %>% 
  filter(uitzending == over_10_dagen & tech_send)

saveRDS(mailen_dag10_tech, file = paste0(config$cz_rds_store, "mailen_dag10_tech.RDS"))
