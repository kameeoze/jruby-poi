module POI
  AREA_REF = org.apache.poi.ss.util.AreaReference
  CELL_REF = org.apache.poi.ss.util.CellReference
  
  def self.Facade(delegate, java_class)
    cls = Class.new
    java_class.java_class.java_instance_methods.select{|e| e.public?}.each do | method |
      args = method.arity.times.collect{|i| "arg#{i}"}.join(", ")
      method_name = method.name.gsub(/([A-Z])/){|e| "_#{e.downcase}"}
      code = "def #{method_name}(#{args}); #{delegate}.#{method.name}(#{args}); end"
      cls.class_eval(code, __FILE__, __LINE__)
    end
    cls
  end
end

require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'area')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'named_range')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'workbook')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'worksheet')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'row')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'cell')
