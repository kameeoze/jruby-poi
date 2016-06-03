# -*- coding: utf-8 -*-
module POI
  class TableRows
    def initialize(document, table)
      @document = document
      @table = table
      @rows = {}
    end
    
    def size
      @table._rows.size
    end
    
    def [](index)
      @rows[index] ||= TableRow.new(@document, @table,  @table._rows[index])
    end
  end
  
  class TableRow < Facade(:poi_row, org.apache.poi.xwpf.usermodel.XWPFTableRow)
    def initialize(document, table, row)
      @document = document
      @table = table
      @row = row
    end
    
    def poi_row
      @row
    end
    
    def table 
      @table
    end
    
    def cells
      @cells ||= TableCells.new(@document, @table, @row)
    end
    
    def [](index)
      cells[index]
    end
    
  end
end