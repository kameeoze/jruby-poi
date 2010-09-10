module POI
  AREA_REF = org.apache.poi.ss.util.AreaReference
  CELL_REF = org.apache.poi.ss.util.CellReference
end

require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'area')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'named_range')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'workbook')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'worksheet')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'row')
require File.join(JRUBY_POI_LIB_PATH, 'poi', 'workbook', 'cell')
