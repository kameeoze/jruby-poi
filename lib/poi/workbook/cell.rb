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
    
    # This is NOT an inexpensive operation. The purpose of this method is merely to get more information
    # out of cell when one thinks the value returned is incorrect. It may have more value in development
    # than in production.
    def error_value
      if poi_cell.getCellType == CELL_TYPE_ERROR
        Java::org.apache.poi.hssf.record.formula.eval.ErrorEval.getText(poi_cell.getErrorCellValue)
      elsif poi_cell.getCellType == CELL_TYPE_FORMULA && poi_cell.getCachedFormulaResultType == CELL_TYPE_ERROR
        formula_evaluator = Java::org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator.new poi_cell.sheet.workbook
        value_of(formula_evaluator.evaluate(poi_cell))
      else
        nil
      end
    end
    
    # returns the formula for this Cell if it has one, otherwise nil
    def formula_value
      poi_cell.getCellType == CELL_TYPE_FORMULA ? poi_cell.getCellFormula : nil
    end
    
    def value
      return nil if poi_cell.nil?
      value_of(cell_value_for_type(poi_cell.getCellType))
    end
    
    def comment
      poi_cell.getCellComment
    end

    def index
      poi_cell.getColumnIndex 
    end    

    # Get the String representation of this Cell's value.
    #
    # If this Cell is a formula you can pass a false to this method and
    # get the formula instead of the String representation.
    def to_s(evaluate_formulas=true)
      return '' if poi_cell.nil?

      if poi_cell.getCellType == CELL_TYPE_FORMULA && evaluate_formulas == false
        formula_value
      else
        value.to_s
      end
    end

    # returns the underlying org.apache.poi.ss.usermodel.Cell
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
          if Java::org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted poi_cell
            Date.parse(Java::org.apache.poi.ss.usermodel.DateUtil.getJavaDate(cell_value.getNumberValue).to_s)
          else
            cell_value.getNumberValue
          end
        when CELL_TYPE_STRING: cell_value.getStringValue
        else
          raise "unhandled cell type[#{cell_value.getCellType}]"
        end
      end
      
      def cell_value_for_type(cell_type)
        return nil if cell_type.nil?
        begin
          case cell_type
          when CELL_TYPE_BLANK: nil
          when CELL_TYPE_BOOLEAN: CELL_VALUE.valueOf(poi_cell.getBooleanCellValue)
          when CELL_TYPE_FORMULA: cell_value_for_type(poi_cell.getCachedFormulaResultType)
          when CELL_TYPE_STRING: CELL_VALUE.new(poi_cell.getStringCellValue)
          when CELL_TYPE_ERROR, CELL_TYPE_NUMERIC: CELL_VALUE.new(poi_cell.getNumericCellValue)
          else
            raise "unhandled cell type[#{poi_cell.getCellType}]"
          end
        rescue
          nil
        end
      end
  end
end

