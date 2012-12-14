JRUBY_POI_LIB_PATH=File.expand_path(File.dirname(__FILE__))

# Java
require 'java'
require File.join(JRUBY_POI_LIB_PATH, 'poi-3.8-20120326.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-ooxml-3.8-20120326.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-ooxml-schemas-3.8-20120326.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-scratchpad-3.8-20120326.jar')
require File.join(JRUBY_POI_LIB_PATH, 'ooxml-lib', 'xmlbeans-2.3.0.jar')
require File.join(JRUBY_POI_LIB_PATH, 'ooxml-lib', 'dom4j-1.6.1.jar')
require File.join(JRUBY_POI_LIB_PATH, 'ooxml-lib', 'stax-api-1.0.1.jar')

# Ruby
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook')
require 'date' # required for Date.parse
