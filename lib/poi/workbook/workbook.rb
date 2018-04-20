require 'tmpdir'
require 'stringio'
require 'java'

module POI
  class Workbook < Facade(:poi_workbook, org.apache.poi.ss.usermodel.Workbook)
    FONT = org.apache.poi.ss.usermodel.Font
    FONT_CONSTANTS = Hash[*FONT.constants.map{|e| [e.downcase.to_sym, FONT.const_get(e)]}.flatten]

    CELL_STYLE = org.apache.poi.ss.usermodel.CellStyle
    CELL_STYLE_CONSTANTS = Hash[*CELL_STYLE.constants.map{|e| [e.downcase.to_sym, CELL_STYLE.const_get(e)]}.flatten]
    
    INDEXED_COLORS = org.apache.poi.ss.usermodel.IndexedColors
    INDEXED_COLORS_CONSTANTS = Hash[*INDEXED_COLORS.constants.map{|e| [e.downcase.to_sym, INDEXED_COLORS.const_get(e)]}.flatten]
    
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
    
    def self.create(filename, options={})
      self.new(filename, nil, options)
    end
    
    attr_reader :filename

    def initialize(filename, io_stream, options={})
      @filename = filename
      @workbook = if io_stream
        org.apache.poi.ss.usermodel.WorkbookFactory.create(io_stream)
      elsif options[:format] == :hssf
        org.apache.poi.hssf.usermodel.HSSFWorkbook.new
      else
        org.apache.poi.xssf.usermodel.XSSFWorkbook.new
      end
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
    
    def create_sheet name='New Sheet'
      # @workbook.createSheet name
      worksheets[name]
    end
    
    def create_style options={}
      font = @workbook.createFont
      set_value( font, :font_height_in_points, options ) do | value |
        value.to_i
      end
      set_value font, :boldweight, options, FONT_CONSTANTS
      set_value font, :color, options, INDEXED_COLORS_CONSTANTS do | value |
        value.index
      end

      style = @workbook.createCellStyle
      [:alignment, :vertical_alignment, :fill_pattern, :border_right, :border_left, :border_top, :border_bottom].each do | sym |
        set_value style, sym, options, CELL_STYLE_CONSTANTS do | value |
          value.to_i
        end
      end
      [:right_border_color, :left_border_color, :top_border_color, :bottom_border_color, :fill_foreground_color, :fill_background_color].each do | sym |
        set_value( style, sym, options, INDEXED_COLORS_CONSTANTS ) do | value |
          value.index
        end
      end
      [:hidden, :locked, :wrap_text].each do | sym |
        set_value style, sym, options
      end
      [:rotation, :indentation].each do | sym |
        set_value( style, sym, options ) do | value |
          value.to_i
        end
      end
      set_value( style, :data_format, options ) do |value|
        @workbook.create_data_format.getFormat(value)
      end
      style.font = font
      style
    end
    
    def set_value on, value_sym, from, using=nil
      return on unless from.has_key?(value_sym)
      value = if using
        using[from[value_sym]]
      else
        from[value_sym]
      end
      value = yield value if block_given?
      on.send("set_#{value_sym}", value)
      on
    end

    def worksheets
      @worksheets ||= Worksheets.new(self)
    end
    
    def named_ranges
      @named_ranges ||= (0...@workbook.number_of_names).collect do | idx |
        NamedRange.new @workbook.get_name_at(idx), self
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
      area = Area.new(reference, self.get_spreadsheet_version)
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
      
      area = AREA_REF.getWholeColumn(get_worksheet_version, area_start, area_end)
      full_ref = "#{area.first_cell.format_as_string}:#{area.last_cell.format_as_string}"
      Area.new(full_ref).in(self)
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

