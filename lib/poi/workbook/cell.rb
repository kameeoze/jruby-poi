module POI
  class Cells
    include Enumerable

    def initialize(row)
      @row = row
      @poi_row = row.poi_row
    end

    def [](index)
      Cell.new(@poi_row.cell(index))
    end

    def size
      @poi_row.physical_number_of_cells
    end

    def each
      it = @poi_row.cell_iterator
      yield Cell.new(it.next) while it.has_next
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
      if poi_cell.cell_type == CELL_TYPE_ERROR
        error_value_from(poi_cell.error_cell_value)
      elsif poi_cell.cell_type == CELL_TYPE_FORMULA && 
            poi_cell.cached_formula_result_type == CELL_TYPE_ERROR
        value_of(formula_evaluator_for(poi_cell.sheet.workbook).evaluate(poi_cell))
      else
        nil
      end
    end
    
    # returns the formula for this Cell if it has one, otherwise nil
    def formula_value
      poi_cell.cell_type == CELL_TYPE_FORMULA ? poi_cell.cell_formula : nil
    end
    
    def value
      return nil if poi_cell.nil?
      value_of(cell_value_for_type(poi_cell.cell_type))
    end
    
    def comment
      poi_cell.cell_comment
    end

    def index
      poi_cell.column_index 
    end    

    # Get the String representation of this Cell's value.
    #
    # If this Cell is a formula you can pass a false to this method and
    # get the formula instead of the String representation.
    def to_s(evaluate_formulas=true)
      return '' if poi_cell.nil?

      if poi_cell.cell_type == CELL_TYPE_FORMULA && evaluate_formulas == false
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
        
        case cell_value.cell_type
        when CELL_TYPE_BLANK: nil
        when CELL_TYPE_BOOLEAN: cell_value.boolean_value
        when CELL_TYPE_ERROR: error_value_from(cell_value.error_value)
        when CELL_TYPE_NUMERIC: numeric_value_from(cell_value)
        when CELL_TYPE_STRING: cell_value.string_value
        else
          raise "unhandled cell type[#{cell_value.cell_type}]"
        end
      end
      
      def cell_value_for_type(cell_type)
        return nil if cell_type.nil?
        begin
          case cell_type
          when CELL_TYPE_BLANK: nil
          when CELL_TYPE_BOOLEAN: CELL_VALUE.value_of(poi_cell.boolean_cell_value)
          when CELL_TYPE_FORMULA: cell_value_for_type(poi_cell.cached_formula_result_type)
          when CELL_TYPE_STRING: CELL_VALUE.new(poi_cell.string_cell_value)
          when CELL_TYPE_ERROR, CELL_TYPE_NUMERIC: CELL_VALUE.new(poi_cell.numeric_cell_value)
          else
            raise "unhandled cell type[#{poi_cell.cell_type}]"
          end
        rescue
          nil
        end
      end
      
      def error_value_from(cell_value)
        # org.apache.poi.hssf.record.formula.eval.ErrorEval.text(cell_value)
        org.apache.poi.ss.usermodel.ErrorConstants.text(cell_value)
      end
      
      def formula_evaluator_for(workbook)
        workbook.creation_helper.create_formula_evaluator
        # org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator.new(workbook)
      end
      
      def numeric_value_from(cell_value)
        if org.apache.poi.ss.usermodel.DateUtil.cell_date_formatted(poi_cell)
          Date.parse(
            org.apache.poi.ss.usermodel.DateUtil.get_java_date(cell_value.number_value).to_s)
        else
          cell_value.number_value
        end
      end
  end
end

