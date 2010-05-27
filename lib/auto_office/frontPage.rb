module AutoOffice

	class FrontPage < AutoOffice::Base
		def initialize
			super('frontpage.application')
			self.show
		end
		
		def hide; end
		def show; end
	end
end