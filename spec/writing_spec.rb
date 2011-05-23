require 'spec_helper'

require 'date'
require 'stringio'

describe "writing Workbooks" do
  it "should create a new empty workbook" do
    name = 'new-workbook.xlsx'
    book = POI::Workbook.create(name)
    book.should_not be_nil
  end
  
  it "should create a new workbook and write something to it" do
    name = TestDataFile.expand_path("timesheet-#{Time.now.strftime('%Y%m%d%H%M%S%s')}.xlsx")
    create_timesheet_spreadsheet(name)
    book = POI::Workbook.open(name)
    book.worksheets.size.should == 1
    book.worksheets[0].name.should == 'Timesheet'
    book.filename.should == name
    book['Timesheet!A3'].should == 'Yegor Kozlov'
    book.cell('Timesheet!J13').formula_value.should == 'SUM(J3:J12)'
    FileUtils.rm_f name
  end
  
  def create_timesheet_spreadsheet name='spec/data/timesheet.xlsx'
    titles = ["Person",	"ID", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Total\nHrs", "Overtime\nHrs", "Regular\nHrs"]
    sample_data = [
      ["Yegor Kozlov", "YK", 5.0, 8.0, 10.0, 5.0, 5.0, 7.0, 6.0],
      ["Gisella Bronzetti", "GB", 4.0, 3.0, 1.0, 3.5, nil, nil, 4.0] 
    ]
    
    book = POI::Workbook.create(name)
    title_style = book.create_style :font_height_in_points => 18, :boldweight => :boldweight_bold,
                                    :alignment => :align_center, :vertical_alignment => :vertical_center
    header_style = book.create_style :font_height_in_points => 11, :color => :white, :fill_foreground_color => :grey_50_percent,
                                     :fill_pattern => :solid_foreground, :alignment => :align_center, :vertical_alignment => :vertical_center
    cell_style  = book.create_style :alignment => :align_center, :border_bottom => :border_thin, :border_top => :border_thin,
                                    :border_left => :border_thin, :border_right => :border_thin, :bottom_border_color => :black,
                                    :right_border_color => :black, :left_border_color => :black, :top_border_color => :black
    form1_style = book.create_style :data_format => '0.00', :fill_pattern => :solid_foreground, :fill_foreground_color => :grey_25_percent,
                                    :alignment => :align_center, :vertical_alignment => :vertical_center
    form2_style = book.create_style :data_format => '0.00', :fill_pattern => :solid_foreground, :fill_foreground_color => :grey_40_percent,
                                    :alignment => :align_center, :vertical_alignment => :vertical_center

    sheet = book.create_sheet 'Timesheet'
    print_setup = sheet.print_setup
    print_setup.landscape = true
    sheet.fit_to_page = true
    sheet.horizontally_center = true
    
    title_row = sheet.rows[0]
    title_row.height_in_points = 45
    title_cell = title_row.cells[0]
    title_cell.value = 'Weekly Timesheet'
    title_cell.style = title_style
    sheet.add_merged_region org.apache.poi.ss.util.CellRangeAddress.valueOf("$A$1:$L$1")
    
    header_row = sheet[1]
    header_row.height_in_points = 40
    titles.each_with_index do | title, index |
      header_cell = header_row[index]
      header_cell.value = title
      header_cell.style = header_style
    end
    
    row_num = 2
    10.times do
      row = sheet[row_num]
      row_num += 1
      titles.each_with_index do | title, index |
        cell = row[index]
        if index == 9
          cell.formula = "SUM(C#{row_num}:I#{row_num})"
          cell.style = form1_style
        elsif index == 11
          cell.formula = "J#{row_num} - K#{row_num}"
          cell.style = form1_style
        else
          cell.style = cell_style
        end
      end
    end
    
    # row with totals below
    sum_row = sheet[row_num]
    row_num += 1
    sum_row.height_in_points = 35
    cell = sum_row[0]
    cell.style = form1_style
    cell = sum_row[1]
    cell.style = form1_style
    cell.value = 'Total Hrs:'
    (2...12).each do | cell_index |
      cell = sum_row[cell_index]
      column = (?A + cell_index).chr
      cell.formula = "SUM(#{column}3:#{column}12)"
      if cell_index > 9
        cell.style = form2_style
      else
        cell.style = form1_style
      end
    end
    row_num += 1
    sum_row = sheet[row_num]
    row_num += 1
    sum_row.height_in_points = 25
    cell = sum_row[0]
    cell.value = 'Total Regular Hours'
    cell.style = form1_style
    cell = sum_row[1]
    cell.formula = 'L13'
    cell.style = form2_style
    sum_row = sheet[row_num]
    row_num += 1
    cell = sum_row[0]
    cell.value = 'Total Overtime Hours'
    cell.style = form1_style
    cell = sum_row[1]
    cell.formula = 'K13'
    cell.style = form2_style

    # set sample data
    sample_data.each_with_index do |each, row_index|
      row = sheet[2 + row_index]
      each.each_with_index do | data, cell_index |
        data = sample_data[row_index][cell_index]
        next unless data
        if data.kind_of? String
          row[cell_index].value = data #.to_java(:string)
        else
          row[cell_index].value = data #.to_java(:double)
        end
      end
    end
    
    # finally set column widths, the width is measured in units of 1/256th of a character width
    sheet.set_column_width 0, 30*256 # 30 characters wide
    (2..9).to_a.each do | column |
      sheet.set_column_width column, 6*256 # 6 characters wide
    end
    sheet.set_column_width 10, 10*256 # 10 characters wide

    book.save
    File.exist?(name).should == true
  end
end
