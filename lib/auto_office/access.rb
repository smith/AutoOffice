module AutoOffice
	class Access < AutoOffice::Base
		def initialize(file_name = "")
			super('access.application')
			open(file_name) if(file_name != "")	
			show
		end
		
		def open(file_name)
			openCurrentDataBase(file_name)
		end

		def run_macro(macro_name)
			doCmd.runMacro(macro_name)
		end
	end
end