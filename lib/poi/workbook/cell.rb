# -*- coding: utf-8 -*-
require 'date'
module POI
  class Cells
    include Enumerable

    def initialize(row)
      @row     = row
      @poi_row = row.poi_row
      @cells   = {}
    end

    def [](index)
      @cells[index] ||= Cell.new(@poi_row.get_cell(index) || @poi_row.create_cell(index), @row)
    end

    def size
      @poi_row.physical_number_of_cells
    end

    def each
      it = @poi_row.cell_iterator
      yield Cell.new(it.next, @row) while it.has_next
    end
  end

  class Cell < Facade(:poi_cell, org.apache.poi.ss.usermodel.Cell)
    DATE_UTIL         = Java::org.apache.poi.ss.usermodel.DateUtil
    CELL              = Java::org.apache.poi.ss.usermodel.Cell
    CELL_VALUE        = Java::org.apache.poi.ss.usermodel.CellValue
    CELL_TYPE_BLANK   = CELL::CELL_TYPE_BLANK
    CELL_TYPE_BOOLEAN = CELL::CELL_TYPE_BOOLEAN
    CELL_TYPE_ERROR   = CELL::CELL_TYPE_ERROR
    CELL_TYPE_FORMULA = CELL::CELL_TYPE_FORMULA
    CELL_TYPE_NUMERIC = CELL::CELL_TYPE_NUMERIC
    CELL_TYPE_STRING  = CELL::CELL_TYPE_STRING
    
    def initialize(cell, row)
      @cell = cell
      @row  = row
    end
    
    def <=> other
      return 1 if other.nil?
      return self.index <=> other.index
    end
    
    # This is NOT an inexpensive operation. The purpose of this method is merely to get more information
    # out of cell when one thinks the value returned is incorrect. It may have more value in development
    # than in production.
    def error_value
      if poi_cell.cell_type == CELL_TYPE_ERROR
        error_value_from(poi_cell.error_cell_value)
      elsif poi_cell.cell_type == CELL_TYPE_FORMULA && 
            poi_cell.cached_formula_result_type == CELL_TYPE_ERROR
            
        cell_value = formula_evaluator.evaluate(poi_cell)
        cell_value && error_value_from(cell_value.error_value)
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
      cast_value
    end

    def formula= new_value
      poi_cell.cell_formula = new_value
      @row.worksheet.workbook.on_formula_update self
      self
    end

    def formula
      poi_cell.cell_formula
    end

    def value= new_value
      set_cell_value new_value
      if new_value.nil?
        @row.worksheet.workbook.on_delete self
      else
        @row.worksheet.workbook.on_update self
      end
      self
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
    
    # :cell_style= comes from the Façade superclass
    alias :style= :cell_style=
    
    def style! options
      self.style = @row.worksheet.workbook.create_style(options)
    end
    
    private
      def cast_value(type = cell_type)
        case type
        when CELL_TYPE_BLANK   then nil
        when CELL_TYPE_BOOLEAN then get_boolean_cell_value
        when CELL_TYPE_ERROR   then nil
        when CELL_TYPE_FORMULA then cast_value(poi_cell.cached_formula_result_type)
        when CELL_TYPE_STRING  then get_string_cell_value
        when CELL_TYPE_NUMERIC
          if DATE_UTIL.cell_date_formatted(poi_cell)
            DateTime.parse(get_date_cell_value.to_s)
          else
            get_numeric_cell_value
          end
        else
          raise "unhandled cell type[#{type}]"
        end
      end

      def workbook
        @row.worksheet.workbook
      end

      def formula_evaluator
        workbook.formula_evaluator
      end

      def error_value_from(cell_value)
        #org.apache.poi.ss.usermodel.ErrorConstants.get_text(cell_value)
        org.apache.poi.ss.usermodel.FormulaError.forInt(cell_value).getString
      end
  end
end

