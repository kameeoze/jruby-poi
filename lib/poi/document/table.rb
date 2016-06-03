# -*- coding: utf-8 -*-
module POI
  class Tables
    def initialize(document)
      @document = document
      @poi_document = document.poi_document
      @tables = {}
    end
    
    def [](index)
      @tables[index] ||= Table.new(@document, @poi_document.tables[index])
    end
    
    def size
      @poi_document.tables.size
    end
  end
  
  class Table < Facade(:poi_table, org.apache.poi.xwpf.usermodel.XWPFTable)
    def initialize(document, table)
      @document = document
      @table = table
    end
    
    alias :_rows :rows
    def rows
      @rows ||= TableRows.new(@document, self)
    end
    
    def poi_table
      @table
    end
  end
end