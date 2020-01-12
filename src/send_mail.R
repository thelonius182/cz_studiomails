library(gmailr)

gm_auth_configure(path = "cz-studiomails.json")
gm_auth(email = "cz.teamservice@gmail.com")


send_reminder <- function(pt_mail, reminder_txt) {
  cz_msg <- gm_mime(To = pt_mail,
                    From = "cz.teamservice@gmail.com",
                    Subject = "Studio Concertzender",
                    body = reminder_txt)
  
  gm_send_message(cz_msg)
}

reminder_text <- function(pres, wanneer, tech, tlf, uitz, hoe_laat) {
  result <- paste0("Hallo ", pres, ",\n\n", 
                  wanneer, ", op ", uitz, 
                  ", heb je studiodienst\n",
                  "met ", tech, " (", tlf, "), van ", hoe_laat, 
                  ".\n\nGraag tot dan!\n\n\n",
                  "Groeten,\nCZ-teamService\n\n")
}