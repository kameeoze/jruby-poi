module POI
  class Area
    def initialize reference
      @ref = reference
    end
    
    def in workbook
      begin
        area.getAllReferencedCells.collect{|c| workbook.cell c.formatAsString}
      rescue
        []
      end
    end
    
    def single_cell_reference?
      @ref == area.getFirstCell.formatAsString
    end
    
    private
      def area
        @area ||= org.apache.poi.ss.util.AreaReference.new(@ref)
      end
  end
end