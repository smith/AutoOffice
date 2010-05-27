require 'win32ole'

module AutoOffice
  class Base < WIN32OLE
  # Need to add CONST module
    OlFolderInbox = 6
    OlMailitem = 0
    XlHide = 3
    XlCSV = 6
    OfficeApps = ['msaccess', 'excel', 'frontpg', 'outlook',
      'powerpnt', 'mspub', 'winword']

    def show
      self.visible = true
    end

    def hide
      self.visible = false
    end

    def quit
      self.quit
    end
  end
end

Dir[File.join(File.dirname(__FILE__), 'auto_office/**/*.rb')].sort.each { |lib| require lib }
