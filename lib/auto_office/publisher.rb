module AutoOffice
	class Publisher < AutoOffice::Base
		def initialize(file_name = "")
			super('publisher.application')
			open(file_name) if(file_name != "")
			show
		end
		
		def hide
			activewindow.visible = false
		end
		
		def show
			activewindow.visible = true
		end
	end
end