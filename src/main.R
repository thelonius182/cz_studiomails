library(googledrive)
library(keyring)
library(readxl)
library(readr)
library(yaml)
library(stringr)
library(dplyr)
library(lubridate)
library(gmailr)
library(purrr)

gm_auth_configure(path = "cz-studiomails.json")
gm_auth(email = "cz.teamservice@gmail.com")

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

# aanmelden bij GD loopt via de procedure die beschreven is in 
# "Operation Missa > Voortgang > 4.Toegangsrechten GD".

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

# teamleden: vooraf opgehaald uit SalesForce en beschikbaar in Downloads
path_teamleden <- paste0(config$gs_downloads, "/studioteamleden.csv")
tbl_teamleden <- read_csv(file = path_teamleden,
                          col_types = cols(NPE01__ALTERNATEEMAIL__C = col_skip())
) %>% 
  mutate(tlf = if_else(!is.na(MOBILEPHONE), MOBILEPHONE, PHONE)) %>% 
  select(-MOBILEPHONE, -PHONE)

tbl_presentatie <- tbl_presentatie.I %>% 
  left_join(tbl_teamleden, by = c("pres" = "NAME")) %>% 
  left_join(tbl_teamleden, by = c("tech" = "NAME")) %>% 
  rename(pres_roep = FIRSTNAME.x,
         pres_mail = EMAIL.x,
         pres_tlf = tlf.x,
         pres_send = TEAMSERVICE_ONTVANGERS__C.x,
         tech_roep = FIRSTNAME.y,
         tech_mail = EMAIL.y,
         tech_tlf = tlf.y,
         tech_send = TEAMSERVICE_ONTVANGERS__C.y
  ) %>% 
  filter(live & (pres_send | tech_send)) 

rm(tbl_presentatie.I, tbl_raw_presentatie, tbl_teamleden)

# overmorgen ----
overmorgen <- today(tzone = "Europe/Amsterdam") + ddays(2L)

mailen_dag2_pres <- tbl_presentatie %>% 
  select(pres, pres_mail, pres_roep, uitzending, pres_send, uren, tech, tech_tlf) %>% 
  filter(uitzending == overmorgen & pres_send)

if (nrow(mailen_dag2_pres) > 0) {
  mail_body_dag2_pres <- "Dag %s,

Overmorgen, %s van %s, ben je weer in de ether. 
Je technicus is dan %s (%s).

Met de groeten van 
CZ-teamservice
" 

  tbl_msg_data <- mailen_dag2_pres %>% 
    mutate(
      To = sprintf("%s <%s>", pres, pres_mail),
      #To = sprintf("%s <lcavdakker@gmail.com>", pres),
      From = "cz.teamservice@gmail.com",
      Subject = sprintf("Herinnering: %s", 
                        str_replace(format(uitzending, "%A %d %B"), " 0", " ")),
      body = sprintf(mail_body_dag2_pres,
                     pres_roep,
                     str_replace(format(uitzending, "%A %d %B"), " 0", " "),
                     uren,
                     tech,
                     tech_tlf)) %>% 
    select(To, From, Subject, body)
  
  write_csv(tbl_msg_data, "tbl-msg-data.csv")
  
  cz_msgs <- tbl_msg_data %>% pmap(gm_mime)
  
  safe_send_msg <- safely(gm_send_message)
  sent_mail <- cz_msgs %>% map(safe_send_msg)
}

mailen_dag2_tech <- tbl_presentatie %>% 
  select(tech, tech_mail, tech_roep, uitzending, tech_send, uren, pres, pres_tlf) %>% 
  filter(uitzending == overmorgen & tech_send)

if (nrow(mailen_dag2_tech) > 0) {
  
  mail_body_dag2_tech <- "Dag %s,
  
Overmorgen, %s van %s, ben je weer in de ether. 
Je presentator is dan %s (%s).
  
Met de groeten van 
CZ-teamservice
" 
  
  tbl_msg_data <- mailen_dag2_tech %>% 
    mutate(
      To = sprintf("%s <%s>", tech, tech_mail),
      # To = sprintf("%s <lcavdakker@gmail.com>", tech),
      From = "cz.teamservice@gmail.com",
      Subject = sprintf("Herinnering: %s", 
                        str_replace(format(uitzending, "%A %d %B"), " 0", " ")),
      body = sprintf(mail_body_dag2_tech,
                     tech_roep,
                     str_replace(format(uitzending, "%A %d %B"), " 0", " "),
                     uren,
                     pres,
                     pres_tlf)) %>% 
    select(To, From, Subject, body)
  
  write_csv(tbl_msg_data, "tbl-msg-data.csv")
  
  cz_msgs <- tbl_msg_data %>% pmap(gm_mime)
  
  safe_send_msg <- safely(gm_send_message)
  sent_mail <- cz_msgs %>% map(safe_send_msg)
}

# over 10 dagen ----
over_10_dagen <- today(tzone = "Europe/Amsterdam") + ddays(10L)

mailen_dag10_pres <- tbl_presentatie %>% 
  select(pres, pres_mail, pres_roep, uitzending, pres_send, uren, tech, tech_tlf) %>% 
  filter(uitzending == over_10_dagen & pres_send)

if (nrow(mailen_dag10_pres) > 0) {
  mail_body_dag10_pres <- "Dag %s,
  
Over 10 dagen, op %s van %s, ben je weer in de ether. 
Je technicus is dan %s (%s).
  
Met de groeten van 
CZ-teamservice
" 
  
  tbl_msg_data <- mailen_dag10_pres %>% 
    mutate(
      To = sprintf("%s <%s>", pres, pres_mail),
      # To = sprintf("%s <lcavdakker@gmail.com>", pres),
      From = "cz.teamservice@gmail.com",
      Subject = sprintf("Herinnering: %s", 
                        str_replace(format(uitzending, "%A %d %B"), " 0", " ")),
      body = sprintf(mail_body_dag10_pres,
                     pres_roep,
                     str_replace(format(uitzending, "%A %d %B"), " 0", " "),
                     uren,
                     tech,
                     tech_tlf)) %>% 
    select(To, From, Subject, body)
  
  write_csv(tbl_msg_data, "tbl-msg-data.csv")
  
  cz_msgs <- tbl_msg_data %>% pmap(gm_mime)
  
  safe_send_msg <- safely(gm_send_message)
  sent_mail <- cz_msgs %>% map(safe_send_msg)
}

mailen_dag10_tech <- tbl_presentatie %>% 
  select(tech, tech_mail, tech_roep, uitzending, tech_send, uren, pres, pres_tlf) %>% 
  filter(uitzending == over_10_dagen & tech_send)

if (nrow(mailen_dag10_tech) > 0) {
  
  mail_body_dag10_tech <- "Dag %s,
  
Over 10 dagen, op %s van %s, ben je weer in de ether. 
Je presentator is dan %s (%s).
  
Met de groeten van 
CZ-teamservice
" 
  
  tbl_msg_data <- mailen_dag10_tech %>% 
    mutate(
      To = sprintf("%s <%s>", tech, tech_mail),
      # To = sprintf("%s <lcavdakker@gmail.com>", tech),
      From = "cz.teamservice@gmail.com",
      Subject = sprintf("Herinnering: %s", 
                        str_replace(format(uitzending, "%A %d %B"), " 0", " ")),
      body = sprintf(mail_body_dag10_tech,
                     tech_roep,
                     str_replace(format(uitzending, "%A %d %B"), " 0", " "),
                     uren,
                     pres,
                     pres_tlf)) %>% 
    select(To, From, Subject, body)
  
  write_csv(tbl_msg_data, "tbl-msg-data.csv")
  
  cz_msgs <- tbl_msg_data %>% pmap(gm_mime)
  
  safe_send_msg <- safely(gm_send_message)
  sent_mail <- cz_msgs %>% map(safe_send_msg)
}
