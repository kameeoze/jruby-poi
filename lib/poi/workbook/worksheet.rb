module POI
  class Worksheets
    include Enumerable
    
    def initialize(workbook)
      @workbook = workbook
      @poi_workbook = workbook.poi_workbook
    end

    def [](index_or_name)
      worksheet = case
        when index_or_name.kind_of?(Numeric)
          @poi_workbook.get_sheet_at(index_or_name) || @poi_workbook.create_sheet
        else 
          @poi_workbook.get_sheet(index_or_name) || @poi_workbook.create_sheet(index_or_name)
      end
      Worksheet.new(worksheet, @workbook)
    end

    def size
      @poi_workbook.number_of_sheets
    end

    def each
      (0...size).each { |i| yield Worksheet.new(@poi_workbook.get_sheet_at(i), @workbook) }
    end
  end

  class Worksheet < Facade(:poi_worksheet, org.apache.poi.ss.usermodel.Sheet)

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
    # row_index as Fixnum: returns the 0-based row
    #
    # row_index as String: assumes a cell reference within this sheet
    #    if the value of the reference is non-nil the value is returned,
    #    otherwise the referenced cell is returned
    def [](row_index)
      if Fixnum === row_index
        rows[row_index]
      else
        ref = org.apache.poi.ss.util.CellReference.new(row_index)
        cell = rows[ref.row][ref.col]
        cell && cell.value ? cell.value : cell
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

