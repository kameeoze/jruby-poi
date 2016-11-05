JRUBY_POI_LIB_PATH=File.expand_path(File.dirname(__FILE__))

require File.join(JRUBY_POI_LIB_PATH, 'poi', 'version')

# Java
require 'java'
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'lib', 'commons-collections4-4.1.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'poi-3.15.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'poi-ooxml-3.15.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'poi-ooxml-schemas-3.15.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'poi-scratchpad-3.15.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'ooxml-lib', 'xmlbeans-2.6.0.jar')
require File.join(JRUBY_POI_LIB_PATH, 'poi-jars', 'ooxml-lib', 'curvesapi-1.04.jar')

#commons-codec-1.10.jar  commons-collections4-4.1.jar  commons-logging-1.2.jar  junit-4.12.jar  log4j-1.2.17.jar

# Ruby
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'version')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook')
require 'date' # required for Date.parse
