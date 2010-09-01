class NamedRange
  
  # takes an instance of org.apache.poi.ss.usermodel.Name, and a POI::Workbook
  def initialize name, workbook
    @name = name
    @workbook = workbook
  end
  
  def name
    @name.getNameName
  end
  
  def sheet
    @workbook[@name.getSheetName]
  end
  
  def formula
    @name.getRefersToFormula
  end
  
  def cells
    area_range? ? all_cells_from_area : @workbook.cell(formula)
  end
  
  def values
    cells.collect{|c| c.value}
  end
  
  private
    def area_range?
      area.nil? == false
    end
    
    def area
      @area ||= begin
        org.apache.poi.ss.util.AreaReference.new formula
      rescue
        nil
      end
    end
    
    def all_cells_from_area
      area.getAllReferencedCells.collect{|c| @workbook.cell c.formatAsString}
    end
end