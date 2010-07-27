require File.join(File.dirname(__FILE__), 'spec_helper')

describe POI::Workbook do
  it "should read an xlsx file" do 
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    book = nil
    lambda { book = POI::Workbook.open(name) }.should_not raise_exception
    book.should be_kind_of POI::Workbook
  end

  it "should read an xls file" do 
    name = TestDataFile.expand_path("simple_with_picture.xls")
    book = nil
    lambda { book = POI::Workbook.open(name) }.should_not raise_exception
    book.should be_kind_of POI::Workbook
  end

  it "should read an ods file" do 
    pending "get a valid, non excel produced ods file" do
      name = TestDataFile.expand_path("simple_with_picture.ods")
      book = nil
      lambda { book = POI::Workbook.open(name) }.should_not raise_exception
      book.should be_kind_of POI::Workbook
    end
  end

  it "should write an open workbook" do
    name = TestDataFile.expand_path("simple_with_picture.xlsx")
    POI::Workbook.open(name) do |book|
      lambda { book.save }.should_not raise_exception
    end
  end

  it "should write an open workbook to a new file" do
    pending "write to a temp directory to avoid git polution" do
      name = TestDataFile.expand_path("simple_with_picture.xlsx")
      new_name = TestDataFile.expand_path("saved-as.xlsx")

      POI::Workbook.open(name) do |book|
        lambda { book.save_as(new_name) }.should_not raise_exception
      end
      
      book = nil
      lambda { book = POI::Workbook.open(new_name) }.should_not raise_exception
      book.should be_kind_of POI::Workbook
      
      File.exists?(new_name).should == true
    end
  end
end
