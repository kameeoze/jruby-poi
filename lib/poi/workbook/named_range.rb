module POI
  class NamedRange
  
    # takes an instance of org.apache.poi.ss.usermodel.Name, and a POI::Workbook
    def initialize name, workbook
      @name = name
      @workbook = workbook
    end
  
    def name
      @name.name_name
    end
  
    def sheet
      @workbook[@name.sheet_name]
    end
  
    def formula
      @name.refers_to_formula
    end
  
    def cells
      [@workbook.cell(formula)].flatten
    end
  
    def values
      cells.collect{|c| c.value}
    end
  end
end