module AutoOffice
	class PowerPoint < AutoOffice::Base
		def initialize(file_name = "")
			super('powerpoint.application')
			open(file_name) if(file_name != "")
			show
		end
	end
end