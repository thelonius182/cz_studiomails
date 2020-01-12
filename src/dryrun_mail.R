library(gmailr)

gm_auth_configure(path = "cz-studiomails.json")
gm_auth(email = "cz.teamservice@gmail.com")


send_reminder <- function(pt_mail, reminder_txt) {
  cz_msg <- gm_mime(To = pt_mail,
                    From = "cz.teamservice@gmail.com",
                    Subject = "Concertzender - Live uitzending",
                    body = reminder_txt)
  
  gm_send_message(cz_msg)
}
