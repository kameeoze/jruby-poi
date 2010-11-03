require 'tmpdir'
require 'stringio'

module POI
  class Workbook
    def self.open(filename_or_stream)
      name, stream = if filename_or_stream.kind_of?(java.io.InputStream)
        [File.join(Dir.tmpdir, "spreadsheet.xlsx"), filename_or_stream]
      elsif filename_or_stream.kind_of?(IO) || StringIO === filename_or_stream || filename_or_stream.respond_to?(:read)
        # NOTE: the String.unpack here can be very inefficient on large files
        [File.join(Dir.tmpdir, "spreadsheet.xlsx"), java.io.ByteArrayInputStream.new(filename_or_stream.read.unpack('c*').to_java(:byte))]
      else
        raise Exception, "FileNotFound" unless File.exists?( filename_or_stream )
        [filename_or_stream, java.io.FileInputStream.new(filename_or_stream)]
      end
      instance = self.new(name, stream)
      if block_given?
        result = yield instance
        return result 
      end
      instance
    end
    
    attr_reader :filename

    def initialize(filename, io_stream)
      @filename = filename
      @workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(io_stream)
    end
    
    def formula_evaluator
      @formula_evaluator ||= @workbook.creation_helper.create_formula_evaluator
    end

    def save
      save_as(@filename)
    end

    def save_as(filename)
      output = output_stream filename
      begin
        @workbook.write(output)
      ensure
        output.close
      end
    end
    
    def output_stream name
      java.io.FileOutputStream.new(name)
    end

    def close
      #noop
    end

    def worksheets
      @worksheets ||= Worksheets.new(self)
    end
    
    def named_ranges
      @named_ranges ||= (0...@workbook.number_of_names).collect do | idx |
        NamedRange.new @workbook.name_at(idx), self
      end
    end

    # reference can be a Fixnum, referring to the 0-based sheet or
    # a String which is the sheet name or a cell reference.
    # 
    # If a cell reference is passed the value of that cell is returned.
    #
    # If the reference refers to a contiguous range of cells an Array of values will be returned.
    #
    # If the reference refers to a multiple columns a Hash of values will be returned by column name.
    def [](reference)
      if Fixnum === reference
        return worksheets[reference]
      end
      
      if sheet = worksheets.detect{|e| e.name == reference}
        return sheet.poi_worksheet.nil? ? nil : sheet
      end

      cell = cell(reference)
      if Array === cell
        cell.collect{|e| e.value}
      elsif Hash === cell
        values = {}
        cell.each_pair{|column_name, cells| values[column_name] = cells.collect{|e| e.value}}
        values
      else
        cell.value
      end
    end

    # takes a String in the form of a 3D cell reference and returns the Cell (eg. "Sheet 1!A1")
    #
    # If the reference refers to a contiguous range of cells an array of Cells will be returned
    def cell reference
      # if the reference is to a named range of cells, get that range and return it
      if named_range = named_ranges.detect{|e| e.name == reference}
        cells = named_range.cells.compact
        if cells.empty?
          return nil
        else
          return cells.length == 1 ? cells.first : cells
        end
      end
      
      # check if the named_range is a full column reference
      if column_reference?(named_range)
        return all_cells_in_column named_range.formula
      end
      
      # if the reference is to an area of cells, get all the cells in that area and return them
      cells = cells_in_area(reference)
      unless cells.empty?
        return cells.length == 1 ? cells.first : cells
      end
      
      if column_reference?(reference)
        return all_cells_in_column reference
      end

      ref = POI::CELL_REF.new(reference)
      single_cell ref
    end
    
    # ref is a POI::CELL_REF instance
    def single_cell ref
      if ref.sheet_name.nil?
        raise 'cell references at the workbook level must include a sheet reference (eg. Sheet1!A1)'
      else
        worksheets[ref.sheet_name][ref.row][ref.col]
      end
    end
    
    def cells_in_area reference
      area = Area.new(reference)
      area.in(self)
    end

    def poi_workbook
      @workbook
    end
    
    def on_update cell
      #clear_all_formula_results
      #formula_evaluator.notify_update_cell cell.poi_cell
    end
    
    def on_formula_update cell
      #clear_all_formula_results
      formula_evaluator.notify_set_formula cell.poi_cell
      formula_evaluator.evaluate_formula_cell(cell.poi_cell)
    end
    
    def on_delete cell
      #clear_all_formula_results
      formula_evaluator.notify_delete_cell cell.poi_cell
    end
    
    def clear_all_formula_results
      formula_evaluator.clear_all_cached_result_values
    end
    
    def all_cells_in_column reference
      sheet_parts = reference.split('!')
      area_parts  = sheet_parts.last.split(':')
      area_start  = "#{sheet_parts.first}!#{area_parts.first}"
      area_end    = area_parts.last
      
      area = org.apache.poi.ss.util.AreaReference.getWholeColumn(area_start, area_end)
      full_ref = "#{area.first_cell.format_as_string}:#{area.last_cell.format_as_string}"
      Area.new(full_ref).in(self)
      # cell_reference = org.apache.poi.ss.util.CellReference.new( reference + "1" )
      # column         = cell_reference.get_col
      # sheet          = cell_reference.get_sheet_name
      # worksheets[sheet].rows.collect{|row| row[column]}
    end
    
    private
      def column_reference? named_range_or_reference
        return false if named_range_or_reference.nil?
        
        reference = named_range_or_reference
        if NamedRange === named_range_or_reference
          reference = named_range_or_reference.formula
        end
        cell_reference = reference.split('!', 2).last
        beginning, ending = cell_reference.split(':')
        !(beginning =~ /\d/ || (ending.nil? ? false : ending =~ /\d/))
      end
  end
end

