module AutoOffice


	# Only works with MAPI (exchange) folders, I think.
	class Outlook < AutoOffice::Base
		def initialize
			super('outlook.application')
			@name_space = getNameSpace('MAPI')
			@inbox = @name_space.getDefaultFolder(OlFolderInbox)
			show
		end

		def create_item(options)
			msg = CreateItem(OlMailitem)
			msg.Display()
			msg.attachments.add(options["Attachments"]) if(options["Attachments"])
			msg["CC"] = options["CC"]
			msg["BCC"] = options["BCC"]										#
			msg["To"] = options["To"]											#
			msg["Subject"] = options["Subject"]           # FIX THIS
			msg["HTMLBody"] = options["HTMLBody"]
			msg
		end

		def send_item(options)
			create_item(options).send
		end

		def show
			@inbox.display
		end
	end
end