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
    CELL              = Java::org.apache.poi.ss.usermodel.Cell
    CELL_VALUE        = Java::org.apache.poi.ss.usermodel.CellValue
    CELL_TYPE_BLANK   = CELL::CELL_TYPE_BLANK
    CELL_TYPE_BOOLEAN = CELL::CELL_TYPE_BOOLEAN
    CELL_TYPE_ERROR   = CELL::CELL_TYPE_ERROR
    CELL_TYPE_FORMULA = CELL::CELL_TYPE_FORMULA
    CELL_TYPE_NUMERIC = CELL::CELL_TYPE_NUMERIC
    CELL_TYPE_STRING  = CELL::CELL_TYPE_STRING
    
    def initialize(cell)
      @cell = cell
    end
    
    def error_value
      @error_value
    end
    
    def value
      return nil if @cell.nil?
      value_of(cell_value_for_type(@cell.getCellType))
    end
    
    def comment
      @cell.getCellComment
    end

    def index
      @cell.getColumnIndex 
    end    

    def to_s(evaluate_formulas=true)
      if @cell.getCellType == CELL_TYPE_FORMULA
        evaluate_formulas ? value.to_s : @cell.getCellFormula
      else
        value.to_s
      end
    end

    def poi_cell
      @cell
    end
    
    private
      def value_of(cell_value)
        return nil if cell_value.nil?
        
        case cell_value.getCellType
        when CELL_TYPE_BLANK: nil
        when CELL_TYPE_BOOLEAN: cell_value.getBooleanValue
        when CELL_TYPE_ERROR: Java::org.apache.poi.hssf.record.formula.eval.ErrorEval.getText(cell_value.getErrorValue)
        when CELL_TYPE_NUMERIC
          if Java::org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted @cell
            Date.parse(Java::org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell_value.getNumberValue).to_s)
          else
            cell_value.getNumberValue
          end
        when CELL_TYPE_STRING: cell_value.getStringValue
        else
          raise "unhandled cell type[#{@cell.getCellType}]"
        end
      end
      
      def cell_value_for_type(cell_type)
        begin
          case cell_type
          when CELL_TYPE_BLANK: nil
          when CELL_TYPE_BOOLEAN: CELL_VALUE.valueOf(@cell.getBooleanCellValue)
          when CELL_TYPE_FORMULA: cell_value_for_type(@cell.getCachedFormulaResultType)
          when CELL_TYPE_STRING: CELL_VALUE.new(@cell.getStringCellValue)
          when CELL_TYPE_ERROR, CELL_TYPE_NUMERIC: CELL_VALUE.new(@cell.getNumericCellValue)
          else
            raise "unhandled cell type[#{@cell.getCellType}]"
          end
        rescue
          @error_value = $!
          nil
        end
      end
  end
end

