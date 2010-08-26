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
    CELL_TYPE_BLANK   = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_BLANK
    CELL_TYPE_BOOLEAN = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_BOOLEAN
    CELL_TYPE_ERROR   = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_ERROR
    CELL_TYPE_FORMULA = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_FORMULA
    CELL_TYPE_NUMERIC = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_NUMERIC
    CELL_TYPE_STRING  = Java::org.apache.poi.ss.usermodel.Cell::CELL_TYPE_STRING
    
    def initialize(cell)
      @cell = cell
    end
    
    def value
      value_of(@cell.getCellType)
    end
    
    def comment
      @cell.getCellComment
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
    
    private
      def value_of(cell_type)
        require 'rubygems'; require 'ruby-debug'; debugger;
        case cell_type
        when CELL_TYPE_BLANK: ''
        when CELL_TYPE_BOOLEAN: @cell.getBooleanCellValue
        when CELL_TYPE_ERROR: @cell.getErrorCellValue
        when CELL_TYPE_FORMULA
          formula_evaluator = Java::org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator.new @cell.sheet.workbook
          cell_value = formula_evaluator.evaluate @cell
          # finish this
        when CELL_TYPE_NUMERIC
          # hmm... dates come in as numbers also
          # handle like: Date.parse(@cell.getDateCellValue.to_s)
          @cell.getNumericCellValue
        when CELL_TYPE_STRING: @cell.getStringCellValue
        else
          raise "unhandled cell type[#{@cell.getCellType}]"
        end
      end
  end
end

