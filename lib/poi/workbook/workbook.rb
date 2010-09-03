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

    # sheet_index can be a Fixnum, referring to the 0-based sheet or
    # a String which is the sheet name or a cell reference.
    # 
    # If a cell reference is passed the value of that cell is returned.
    def [](reference)
      begin
        cell = cell(reference)
        Array === cell ? cell.collect{|e| e.value} : cell.value
      rescue
        answer = worksheets[reference]
        answer.poi_worksheet.nil? ? nil : answer
      end
    end

    # takes a String in the form of a 3D cell reference and returns the Cell (eg. "Sheet 1!A1")
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
      
      # if the reference is to an area of cells, get all the cells in that area and return them
      cells = cells_in_area(reference)
      unless cells.empty?
        return cells.length == 1 ? cells.first : cells
      end
      
      ref = org.apache.poi.ss.util.CellReference.new(reference)
      if ref.sheet_name.nil?
        raise 'cell references at the workbook level must include a sheet reference (eg. Sheet1!A1)'
      else
        worksheets[ref.sheet_name][ref.row][ref.col]
      end
    end
    
    def cells_in_area reference
      area = Area.new(reference)
      if area.single_cell_reference?
        []
      else
        area.in(self).compact
      end
    end

    def poi_workbook
      @workbook
    end
  end
end

