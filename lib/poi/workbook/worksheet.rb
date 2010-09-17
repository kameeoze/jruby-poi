module POI
  class Worksheets
    include Enumerable
    
    def initialize(workbook)
      @workbook = workbook
      @poi_workbook = workbook.poi_workbook
    end

    def [](index)
      worksheet = case
        when index.kind_of?(Numeric)
          @poi_workbook.sheet_at(index)
        else 
          @poi_workbook.get_sheet(index)
      end
      Worksheet.new(worksheet, @workbook)
    end

    def size
      @poi_workbook.number_of_sheets
    end

    def each
      (0...size).each { |i| yield Worksheet.new(@poi_workbook.sheet_at(i), @workbook) }
    end
  end

  class Worksheet
    def initialize(worksheet, workbook)
      @worksheet = worksheet
      @workbook  = workbook
    end
    
    def name
      @worksheet.sheet_name
    end

    def first_row
      @worksheet.first_row_num
    end

    def last_row
      @worksheet.last_row_num
    end

    def rows
      @rows ||= Rows.new(self)
    end
    
    # Accepts a Fixnum or a String as the row_index
    #
    # row_index as Fixnum - returns the 0-based row
    #
    # row_index as String - assumes a cell reference within this sheet and returns the cell value for that reference
    def [](row_index)
      if Fixnum === row_index
        rows[row_index]
      else
        ref = org.apache.poi.ss.util.CellReference.new(row_index)
        rows[ref.row][ref.col].value
      end
    end

    def poi_worksheet
      @worksheet
    end
    
    def workbook
      @workbook
    end
  end
end

