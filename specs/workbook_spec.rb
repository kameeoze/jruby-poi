require File.join(File.dirname(__FILE__), 'spec_helper')
require 'date'

describe POI::Workbook do
  it "should open a workbook and allow access to its worksheets" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)

    book.worksheets.size.should == 3
  end
end

describe POI::Worksheets do
  it "should allow indexing worksheets by ordinal" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)

    book.worksheets[0].name.should == "text & pic"
    book.worksheets[1].name.should == "numbers"
    book.worksheets[2].name.should == "dates"
  end

  it "should allow indexing worksheets by name" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)

    book.worksheets["text & pic"].name.should == "text & pic"
    book.worksheets["numbers"].name.should == "numbers"
    book.worksheets["dates"].name.should == "dates"
  end

  it "should be enumerable" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)
    book.worksheets.should be_kind_of Enumerable

    book.worksheets.each do |sheet|
      sheet.should be_kind_of POI::Worksheet
    end

    book.worksheets.size.should == 3
    book.worksheets.collect.size.should == 3
  end
end

describe POI::Rows do
  it "should be enumerable" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)
    sheet = book.worksheets["text & pic"]
    sheet.rows.should be_kind_of Enumerable
    
    sheet.rows.each do |row|
      row.should be_kind_of POI::Row
    end

    sheet.rows.size.should == 6
    sheet.rows.collect.size.should == 6
  end
end

describe POI::Cells do
  it "should be enumerable" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)
    sheet = book.worksheets["text & pic"]
    rows = sheet.rows
    cells = rows[0].cells

    cells.should be_kind_of Enumerable
    cells.size.should == 1
    cells.collect.size.should == 1
  end
  
  it "should provide dates for date cells" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)
    sheet = book.worksheets["dates"]
    rows = sheet.rows

    date = Date.parse('2010-02-28')
    (1..16).each do |row|
      (0..2).each do |col|
        rows[row].cells[col].value.should == date
        date = date + 1
      end
    end
    Date.parse("2010-04-12").should == date
  end

  it "should provide the cell value as a string" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = POI::Workbook.open(name)
    sheet = book.worksheets["text & pic"]
    rows = sheet.rows

    rows[0].cells[0].to_s.should == "This"
    rows[1].cells[0].to_s.should == "is"
    rows[2].cells[0].to_s.should == "an"
    rows[3].cells[0].to_s.should == "Excel"
    rows[4].cells[0].to_s.should == "XLSX"
    rows[5].cells[0].to_s.should == "workbook"
  end
end

