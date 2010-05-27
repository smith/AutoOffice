module AutoOffice
	class Word < AutoOffice::Base
		def initialize(file_name = "")
			super('word.application')
			if(file_name == "")
				documents.add
			else
				open(file_name)
			end
			show
		end
		
		def open(file_name)
			documents.open(file_name)
		end
		
		def write(text = "")
			selection.typeText('text' => text)
		end
	end
end

