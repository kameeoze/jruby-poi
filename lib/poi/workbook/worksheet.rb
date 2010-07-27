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
          @poi_workbook.getSheetAt(index)
        else 
          @poi_workbook.getSheet(index)
      end
      Worksheet.new(worksheet)
    end

    def size
      @poi_workbook.getNumberOfSheets
    end

    def each
      (0...size).each { |i| yield Worksheet.new(@poi_workbook.getSheetAt(i)) }
    end
  end

  class Worksheet
    def initialize(worksheet = nil)
      @worksheet = worksheet
    end

    def name
      @worksheet.getSheetName
    end

    def rows
      Rows.new(self)
    end

    def poi_worksheet
      @worksheet
    end
  end
end

