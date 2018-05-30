

# the parts marked with @column_name_from_excel_capitalisation_matters@
# get replaced by the matching column entry
# do not use spaces or punctuation in headings
# assume the To email address heading is "email"
# assume the attachment filename heading is "attachment"
# assume the email subject heading is "subject"
# and excel file is called emails.xlsx and the details are on sheet 1
# in the same folder as this script and the attachment files

# Things you need to do:
# 1. Have the details to be merged in the spreadsheet, one row of headings
# 2. Set the From Address at the Top of this script
# 3. Customise the example message text to be the message you want
# 4. When you run the script, choose the excel file in the Open file dialog

# This is a very first verion that doesn't escape single quotes properly and similar
# but it might be useful to people as is.
# It also doesn't do anything clever about entries in the spreadsheet being left blank

message_text <- "
Dear @Firstname@,

Thank you for your picture of a @animal@.

I am writing to say here is your data attached as a file.

Yours Sincerely,
Whoever
"

# load helper library or install then load helper library
if(!require(readxl, quietly = TRUE, warn.conflicts = FALSE)){
  install.packages("readxl")
  require(readxl, quietly = TRUE, warn.conflicts = FALSE)
}

excel_file_location <- file.choose()
data_file <- normalizePath(excel_file_location)
folder_with_everything <- dirname(excel_file_location)

merge_data <- read_excel(data_file, sheet = 1)
merge_data$attachment[!is.na(merge_data$attachment)] <- paste(folder_with_everything,
                                                              merge_data$attachment[!is.na(merge_data$attachment)],
                                                              sep="/")
headings <- names(merge_data)

# if you want to automatically send, remove the -- from in front of the send msg
applescript_mail.app <- '
tell application "Mail"
activate
set theSubject to "@subject@" -- the subject
set theContent to "@mail_body@" -- the content
set theAddress to "@email@" -- the receiver 
set theAttachmentFile to (POSIX file "@theUnixPath@") as string -- attachment in Mac path format
set msg to make new outgoing message with properties {subject:theSubject, content:theContent, visible:true}
tell msg to make new to recipient at end of every to recipient with properties {address:theAddress}
tell msg to make new attachment with properties {file name:theAttachmentFile as alias}
-- send msg
end tell
'

for (i in 1:nrow(merge_data)){
  message_body <- message_text
  for (j in 1:length(headings)){
    message_body <- gsub(paste0("@", headings[j], "@"), merge_data[i,j], message_body)
  }
  osa_command <- applescript_mail.app
  osa_command <- gsub("@subject@", merge_data$subject[i], osa_command)
  osa_command <- gsub("@mail_body@", message_body, osa_command)
  osa_command <- gsub("@theUnixPath@", merge_data$attachment[i], osa_command)
  osa_command <- gsub("@email@", merge_data$email[i], osa_command)
  osa_command <- paste0("osascript -e '", osa_command, "'")
  system(osa_command)
}
