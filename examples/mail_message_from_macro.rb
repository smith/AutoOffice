# Run a macro in access and send the result as an email
require 'auto_office'

a = Access.new("c:\\documents and settings\\me\\my documents\\db\\callingcards.mdb")
a.run_macro("Macro1")

cards_left = (3000 - a.screen.activeDataSheet.currentRecord).to_s
a.quit

html = "<p style='font-family: arial;font-size:16pt'>"
html << cards_left + "</p>"

mail_item_options = {
	"To" => "Users",
	"Subject" => "Approximate Number of Free CDs Left for new Subscribers",
	"HTMLBody" => html
}

ol = Outlook.new
ol.send_item(mail_item_options)


