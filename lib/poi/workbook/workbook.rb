require 'tmpdir'
require 'stringio'

module POI
  class Workbook
    def self.open(filename_or_stream)
      name, stream = if filename_or_stream.kind_of?(java.io.InputStream)
        [File.join(Dir.tmpdir, "spreadsheet.xlsx"), filename_or_stream]
      elsif filename_or_stream.kind_of?(IO) || StringIO === filename_or_stream || filename_or_stream.respond_to?(:read)
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
      @workbook.write(java.io.FileOutputStream.new(filename))
    end

    def close
      #noop
    end

    def worksheets
      Worksheets.new(self)
    end

    def poi_workbook
      @workbook
    end
  end
end

