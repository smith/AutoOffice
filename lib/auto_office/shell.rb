module AutoOffice
  class Shell < AutoOffice::Base
    def initialize
      super('wscript.shell')
    end
		
    def Shell.kill_office
      system("taskkill" + 
	OfficeApps.map {|app| " /im #{app}.exe"}.to_s + " /f")
    end
			
    def Shell.kill(app)
      system("taskkill /im #{app}.exe /f")
    end
  end
end
              
