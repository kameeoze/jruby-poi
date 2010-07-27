module POI
  class Cells
    include Enumerable

    def initialize(row)
      @row = row
      @poi_row = row.poi_row
    end

    def [](index)
      Cell.new(@poi_row.getCell(index))
    end

    def size
      @poi_row.getPhysicalNumberOfCells
    end

    def each
      it = @poi_row.cellIterator
      yield Cell.new(it.next) while it.hasNext
    end
  end

  class Cell
    def initialize(cell)
      @cell = cell
    end

    def index
      @cell.getColumnIndex 
    end    

    def to_s
      @cell.getStringCellValue
    end

    def poi_cell
      @cell
    end
  end
end

