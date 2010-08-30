require File.join(File.dirname(__FILE__), 'spec_helper')
require 'date'
require 'stringio'

describe POI::Workbook do
  it "should open a workbook and allow access to its worksheets" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    book = POI::Workbook.open(name)
    book.worksheets.size.should == 4
    book.filename.should == name
  end

  it "should be able to create a Workbook from an IO object" do
    content = StringIO.new(open(TestDataFile.expand_path("various_samples.xlsx"), 'rb'){|f| f.read})
    book = POI::Workbook.open(content)
    book.worksheets.size.should == 4
    book.filename.should =~ /spreadsheet.xlsx$/
  end

  it "should be able to create a Workbook from a Java input stream" do
    content = java.io.FileInputStream.new(TestDataFile.expand_path("various_samples.xlsx"))
    book = POI::Workbook.open(content)
    book.worksheets.size.should == 4
    book.filename.should =~ /spreadsheet.xlsx$/
  end
end

describe POI::Worksheets do
  it "should allow indexing worksheets by ordinal" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    book = POI::Workbook.open(name)

    book.worksheets[0].name.should == "text & pic"
    book.worksheets[1].name.should == "numbers"
    book.worksheets[2].name.should == "dates"
    book.worksheets[3].name.should == "bools & errors"
  end

  it "should allow indexing worksheets by name" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    book = POI::Workbook.open(name)

    book.worksheets["text & pic"].name.should == "text & pic"
    book.worksheets["numbers"].name.should == "numbers"
    book.worksheets["dates"].name.should == "dates"
  end

  it "should be enumerable" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    book = POI::Workbook.open(name)
    book.worksheets.should be_kind_of Enumerable

    book.worksheets.each do |sheet|
      sheet.should be_kind_of POI::Worksheet
    end

    book.worksheets.size.should == 4
    book.worksheets.collect.size.should == 4
  end
end

describe POI::Rows do
  it "should be enumerable" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    book = POI::Workbook.open(name)
    sheet = book.worksheets["text & pic"]
    sheet.rows.should be_kind_of Enumerable
    
    sheet.rows.each do |row|
      row.should be_kind_of POI::Row
    end

    sheet.rows.size.should == 7
    sheet.rows.collect.size.should == 7
  end
end

describe POI::Cells do
  before :each do
    @name = TestDataFile.expand_path("various_samples.xlsx")
    @book = POI::Workbook.open(@name)
  end

  def book
    @book
  end

  def name
    @name
  end
  
  it "should be enumerable" do
    sheet = book.worksheets["text & pic"]
    rows = sheet.rows
    cells = rows[0].cells

    cells.should be_kind_of Enumerable
    cells.size.should == 1
    cells.collect.size.should == 1
  end

  it "should provide dates for date cells" do
    sheet = book.worksheets["dates"]
    rows = sheet.rows
    
    dates_by_column = [
      (Date.parse('2010-02-28')..Date.parse('2010-03-14')),
      (Date.parse('2010-03-14')..Date.parse('2010-03-28')),
      (Date.parse('2010-03-28')..Date.parse('2010-04-11'))]
    (0..2).each do |col|
      dates_by_column[col].each_with_index do |date, index|
        row = index + 1
        rows[row][col].value.should equal_at_cell(date, row, col)
      end
    end
  end

  it "should provide numbers for numeric cells" do
    sheet = book.worksheets["numbers"]
    rows = sheet.rows
    
    (1..15).each do |number|
      row = number
      rows[row][0].value.should equal_at_cell(number,            row, 0)
      rows[row][1].value.should equal_at_cell(number ** 2,       row, 1)
      rows[row][2].value.should equal_at_cell(number ** 3,       row, 2)
      rows[row][3].value.should equal_at_cell(Math.sqrt(number), row, 3)
    end
    
    rows[9][0].to_s.should == '9.0'
    rows[9][1].to_s.should == '81.0'
    rows[9][2].to_s.should == '729.0'
    rows[9][3].to_s.should == '3.0'
  end


  it "should handle array access from the workbook down to cells" do
    book[1][9][0].to_s.should == '9.0'
    book[1][9][1].to_s.should == '81.0'
    book[1][9][2].to_s.should == '729.0'
    book[1][9][3].to_s.should == '3.0'

    book["numbers"][9][0].to_s.should == '9.0'
    book["numbers"][9][1].to_s.should == '81.0'
    book["numbers"][9][2].to_s.should == '729.0'
    book["numbers"][9][3].to_s.should == '3.0'
  end

  it "should provide error text for error cells" do
    sheet = book.worksheets["bools & errors"]
    rows = sheet.rows

    rows[6][0].value.should == '~CIRCULAR~REF~'
    rows[7][0].value.should == '#DIV/0!'
    rows[8][0].value.should == '#N/A'
    rows[9][0].value.should == '#NAME?'
    rows[10][0].value.should == '#NULL!'
    rows[11][0].value.should == '#NUM!'
    rows[12][0].value.should == '#REF!'
    rows[13][0].value.should == '#VALUE!'
    lambda{ rows[14][0].value }.should raise_error(Java::java.lang.RuntimeException)
    
    rows[6][0].to_s.should == '~CIRCULAR~REF~'
    rows[7][0].to_s.should == '#DIV/0!'
    rows[8][0].to_s.should == '#N/A'
    rows[9][0].to_s.should == '#NAME?'
    rows[10][0].to_s.should == '#NULL!'
    rows[11][0].to_s.should == '#NUM!'
    rows[12][0].to_s.should == '#REF!'
    rows[13][0].to_s.should == '#VALUE!'
    lambda{ rows[14][0].to_s }.should raise_error(Java::java.lang.RuntimeException)
  end

  it "should provide booleans for boolean cells" do
    sheet = book.worksheets["bools & errors"]
    rows = sheet.rows
    rows[1][0].value.should == false
    rows[1][0].to_s.should == 'false'
    
    rows[1][1].value.should == false
    rows[1][1].to_s.should == 'false'
    
    rows[2][0].value.should == true
    rows[2][0].to_s.should == 'true'
    
    rows[2][1].value.should == true
    rows[2][1].to_s.should == 'true'
  end

  it "should provide the cell value as a string" do
    sheet = book.worksheets["text & pic"]
    rows = sheet.rows

    rows[0][0].value.should == "This"
    rows[1][0].value.should == "is"
    rows[2][0].value.should == "an"
    rows[3][0].value.should == "Excel"
    rows[4][0].value.should == "XLSX"
    rows[5][0].value.should == "workbook"
    rows[9][0].value.should == 'This is an Excel XLSX workbook.'
    

    rows[0][0].to_s.should == "This"
    rows[1][0].to_s.should == "is"
    rows[2][0].to_s.should == "an"
    rows[3][0].to_s.should == "Excel"
    rows[4][0].to_s.should == "XLSX"
    rows[5][0].to_s.should == "workbook"
    rows[9][0].to_s.should == 'This is an Excel XLSX workbook.'
  end
  
  it "should provide formulas instead of string-ified values" do
    sheet = book.worksheets["numbers"]
    rows = sheet.rows

    (1..15).each do |number|
      row = number
      rows[row][0].to_s(false).should == "#{number}.0"
      rows[row][1].to_s(false).should == "A#{row + 1}*A#{row + 1}"
      rows[row][2].to_s(false).should == "B#{row + 1}*A#{row + 1}"
      rows[row][3].to_s(false).should == "SQRT(A#{row + 1})"
    end

    sheet = book.worksheets["bools & errors"]
    rows = sheet.rows
    rows[1][0].to_s(false).should == '1=2'
    rows[1][1].to_s(false).should == 'FALSE'
    rows[2][0].to_s(false).should == '1=1'
    rows[2][1].to_s(false).should == 'TRUE'
    rows[14][0].to_s(false).should == 'foobar(1)'
    
    sheet = book.worksheets["text & pic"]
    sheet.rows[9][0].to_s(false).should == 'CONCATENATE(A1," ", A2," ", A3," ", A4," ", A5," ", A6,".")'
  end
  
  it "should handle getting values out of 'non-existent' cells" do
    sheet = book.worksheets["bools & errors"]
    sheet.rows[14][2].value.should be_nil
  end

end

