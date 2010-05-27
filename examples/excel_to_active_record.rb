# Get data from excel sheets and insert into DB with ActiveRecord
require 'active_record'
require 'auto_office'

# Set ActiveRecord logging options
ActiveRecord::Base.logger = Logger.new(STDERR)
ActiveRecord::Base.colorize_logging = false

# Get Date (will be used later)
print "Enter Date: "
date = gets.chomp!

# Create database in memory
ActiveRecord::Base.establish_connection(
    :adapter => "mysql",
    :host => "hostname",
    :username => "root",
    :password => "password",
		:database => "my_app_development"
	)
    
class Stat < ActiveRecord::Base
end

# Initialize db rows
users = Array.new
users << Stat.new(:submitter => "Bobby Sue" , :agent_id => 1813, :date => date)
users << Stat.new(:submitter => "Mary Lou", :agent_id => 1811, :date => date)
users << Stat.new(:submitter => "Billy Jean", :agent_id => 1821, :date => date)

# Open remedy report
xl = Excel.new(Dir.getwd + "/dtbu.xls")
xl.range("a1:a50").each do |cell|
	users.each do |record|
		if(cell.value == record.submitter)
			row = cell.address.gsub!(/\$A\$/,"")
			record.total_call_log_tickets = xl.range("b" + row).value
		end
	end
end
xl.workbooks.close

# Open Saves/Cancels
xl.open(Dir.getwd + "/sc.xls")
xl.range("a1:a50").each do |cell|
	users.each do |record|
		if(cell.value == record.submitter)
			row = cell.address.gsub!(/\$A\$/,"")
			record.saves = xl.range("e" + row).value
			record.cancels = xl.range("c" + row).value
		end
	end
end
xl.workbooks.close

xl.quit

# Save records
users.each {|record| record.save}

