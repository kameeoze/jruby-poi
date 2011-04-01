module POI
  AREA_REF = org.apache.poi.ss.util.AreaReference
  CELL_REF = org.apache.poi.ss.util.CellReference
  
  def self.Facade(delegate, java_class)
    cls = Class.new
    java_class.java_class.java_instance_methods.select{|e| e.public?}.each do | method |
      args = method.arity.times.collect{|i| "arg#{i}"}.join(", ")
      method_name = method.name.gsub(/([A-Z])/){|e| "_#{e.downcase}"}
      code = "def #{method_name}(#{args}); #{delegate}.#{method.name}(#{args}); end"

      if method_name =~ /^get_[a-z]/
        alias_method_name = method_name.sub('get_', '')
        code += "\nalias :#{alias_method_name} :#{method_name}"
        if method.return_type.to_s == 'boolean'
          code += "\nalias :#{alias_method_name}? :#{method_name}"
        end
      elsif method_name =~ /^set_[a-z]/ && method.arity == 1
        alias_method_name = "#{method_name.sub('set_', '')}"
        code += "\nalias :#{alias_method_name}= :#{method_name}"
        if method.argument_types.first.to_s == 'boolean'
          code += "\ndef #{alias_method_name}!; #{alias_method_name} = true; end"
        end
      elsif method.return_type.to_s == 'boolean' && method_name =~ /is_/
        code += "\nalias :#{method_name.sub('is_', '')}? :#{method_name}"
      elsif method.return_type.nil? && (method.argument_types.nil? || method.argument_types.empty?)
        code += "\nalias :#{method_name}! :#{method_name}"
      end
      
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
