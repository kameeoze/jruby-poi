module POI
  class Rows
    include Enumerable
    
    def initialize(worksheet)
      @worksheet = worksheet
      @poi_worksheet = worksheet.poi_worksheet
    end

    def [](index)
      Row.new(@poi_worksheet.getRow(index))
    end

    def size 
      @poi_worksheet.getPhysicalNumberOfRows 
    end

    def each
      it = @poi_worksheet.rowIterator
      yield Row.new(it.next) while it.hasNext
    end
  end

  class Row
    def initialize(row)
      @row = row
    end

    def cells
      Cells.new(self)
    end

    def index
      @row.getRowNum
    end    

    def poi_row
      @row
    end
  end
end

