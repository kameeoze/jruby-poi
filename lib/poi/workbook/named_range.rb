module POI
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
      [@workbook.cell(formula)].flatten
    end
  
    def values
      cells.collect{|c| c.value}
    end
  end
end