module POI
  class Workbook
    def self.open(filename)
      raise Exception, "FileNotFound" unless File.exists? filename
      instance = self.new(filename)

      if block_given?
        result = yield instance
        return result 
      end

      instance
    end

    def initialize(filename)
      @filename = filename
      @workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(java.io.FileInputStream.new(filename))
    end

    def save
      @workbook.write(java.io.FileOutputStream.new(@filename))
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

