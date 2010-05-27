# Send a mail message containing web page content

require 'auto_office'
require 'open-uri'

ol = Outlook.new

print "Enter URL for new article: "
url = gets.chomp!

title = url.gsub(/http:\/\/10\.40\.64\.69\/kb\/index\.php\//, "").gsub("_", " ")

# html holds the email body
# Add email body text
html = "<div style=\"font-family:arial;font-size:10pt\">" +
			 "<p>A new article has been added to the <a href='http://10.40.64.69/'>" +
			 "Knowledge Base</a>:</p><p><a href='#{url}'>#{url}</a></p>" +
			 "<p>Please read and understand. The full article has been reproduced " +
			 "below for your convenience.</p><p>Thanks</p></div>"


# Add page content
inDiv = false
open(url).each_line do |line| 
	if(line.to_s.match("<div id=\"content\">"))
		inDiv = true
		html << line
	elsif(line.to_s.match("<div id=\"column-one\">"))
		inDiv = false
	elsif(inDiv)
		html << line
	end
end
		
mail_item_options = {
	"To" => "Users",
	"Subject" => "New KB Article: #{title}",
	"HTMLBody" => html
}

ol.send_item(mail_item_options)
