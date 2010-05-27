module AutoOffice
  class Excel < AutoOffice::Base
    DefaultRange = 'a1'
		
    def initialize(file_name = "")
      super('excel.application')
      if(file_name == "") 
	workbooks.add
      else
	open(file_name)
      end
      show
    end
		
    def auto_filter(range = DefaultRange)
      activesheet.enableautofilter = true
      activesheet.range(range).autofilter.range = range
    end

    def auto_fit(range = DefaultRange)
      range(range).columns.autofit
    end
		
    def copy(range = DefaultRange)
      range(range).copy
    end


    def hide_rows(range = DefaultRange)
      range(range).rows.hidden = true
    end

    def first_nil(range = DefaultRange)
      num_cells = 0
      range(range).each do |cell|
	num_cells += 1
	break if(cell.value == nil)
      end
      num_cells
    end
		
    def close
      activeworkbook.close
    end
		
    def delete(range = DefaultRange)
      range(range).delete
    end

    def get(range = DefaultRange)
      range(range).value
    end
		
    def hide_pictures
      activeworkbook.displaydrawingobjects = XlHide
    end

    def open(file_name)
      workbooks.open(file_name)
    end

    def paste
      activesheet.paste
    end
		
    def save
      activeworkbook.save
    end

    def save_as_csv(file_name = activeworkbook.name.gsub(/\.xls/, ''))
      activeworkbook.saveas(file_name, XlCSV)
    end

    def select(range = DefaultRange)
      range(range).select
    end

    def set(range = DefaultRange, value = nil)
      range(range).value = value
    end	

    def to_percent(range)
      range(range).style = "Percent"
    end

    def find_row(range="A1", value = "")
      row = 1
      range(range).each do |r|
	if(r.value.nil? || !(r.value.match(value)))
	  row += 1
	else
	  break
	end
      end
      row
    end

    def activate_workbook(sheet)
      self.Windows(sheet).Activate()
    end

    def move(sheet)
      activesheet.move("After" => workbooks(sheet).sheets(1))
    end

    def delete_sheet(sheet)
      activeworkbook.sheets(sheet).delete
    end

    # Create a new pivot table. Returns the pivot table data address
    def pivot_table_wizard(options)
      pivot_cache = activeworkbook.pivotcaches.add(
	"SourceType" => options["SourceType"],
	"SourceData" => range(options["SourceData"])
      )
      pivot_cache.createpivottable(
	"TableDestination" => "",
	"TableName" => options["TableName"],
      )
      cells = activeSheet.cells(3,1)
      activesheet.pivottablewizard(
	"TableDestination" => cells
      )
      cells.select
      pivot_tables = activesheet.pivottables(options["TableName"])
      pivot_tables.AddFields(
	"RowFields" => options["RowFields"],
	"ColumnFields" => options["ColumnFields"]
      )
      pivot_fields = pivot_tables.PivotFields(options["PivotFields"])
      pivot_fields["Orientation"] = 4
      pivot_fields["Caption"] = "Count of " + options["PivotFields"]
      pivot_fields["Function"] = -4112
      pivot_tables.columnrange.address
    end
  end
end
