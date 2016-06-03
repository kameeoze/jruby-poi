# -*- coding: utf-8 -*-
module POI
  class TableCells
    def initialize(document, table, row)
      @document = document
      @table = table
      @row = row
      @cells = {}
    end
    
    def size
      @row.table_cells.size
    end
    
    def [](index)
      @cells[index] ||= TableCell.new(@document, @table, @row.cell(index))
    end
  end
  
  class TableCell < Facade(:poi_cell, org.apache.poi.xwpf.usermodel.XWPFTableCell)
    def initialize(document, table, cell)
      @document = document
      @table = table
      @cell = cell
    end
    
    def poi_cell
      @cell
    end
    
    def table
      @table
    end
  end
end