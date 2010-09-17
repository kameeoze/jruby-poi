module POI
  class Rows
    include Enumerable
    
    def initialize(worksheet)
      @worksheet = worksheet
      @poi_worksheet = worksheet.poi_worksheet
      @rows = {}
    end

    def [](index)
      @rows[index] ||= Row.new(@poi_worksheet.row(index), @worksheet)
    end

    def size 
      @poi_worksheet.physical_number_of_rows 
    end

    def each
      it = @poi_worksheet.row_iterator
      yield Row.new(it.next, @worksheet) while it.has_next
    end
  end

  class Row
    def initialize(row, worksheet)
      @row       = row
      @worksheet = worksheet
    end
    
    def [](index)
      return nil if poi_row.nil?
      cells[index]
    end

    def cells
      @cells ||= Cells.new(self)
    end

    def index
      return nil if poi_row.nil?
      poi_row.row_num
    end    

    def poi_row
      @row
    end
    
    def worksheet
      @worksheet
    end
  end
end

