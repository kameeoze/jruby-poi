require 'spec_helper'

describe POI::Workbook do
  before :each do
    @mock_output_stream = nil
    class POI::Workbook
      def mock_output_stream name
        @mock_output_stream ||= Java::jrubypoi.MockOutputStream.new name
        @mock_output_stream
      end

      alias_method :original_output_stream, :output_stream unless method_defined?(:original_output_stream)
      alias_method :output_stream, :mock_output_stream
    end
  end

  after :each do
    @mock_output_stream = nil
    class POI::Workbook
      alias_method :output_stream, :original_output_stream
    end
  end

  it "should read an xlsx file" do 
    name = TestDataFile.expand_path("various_samples.xlsx")
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

  it "should read an ods file", :unimplemented => true do 
    name = TestDataFile.expand_path("spreadsheet.ods")
    book = nil
    lambda { book = POI::Workbook.open(name) }.should_not raise_exception
    book.should be_kind_of POI::Workbook
  end

  it "should write an open workbook" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    POI::Workbook.open(name) do |book|
      lambda { book.save }.should_not raise_exception
      verify_mock_output_stream book.instance_variable_get("@mock_output_stream"), name
    end
  end

  it "should write an open workbook to a new file" do
    name = TestDataFile.expand_path("various_samples.xlsx")
    new_name = TestDataFile.expand_path("saved-as.xlsx")

    POI::Workbook.open(name) do |book|
      @mock_output_stream.should be_nil
      lambda { book.save_as(new_name) }.should_not raise_exception
      verify_mock_output_stream book.instance_variable_get("@mock_output_stream"), new_name
    end
  end
  
  def verify_mock_output_stream mock_output_stream, name
    mock_output_stream.should_not be_nil
    name.should == mock_output_stream.name
    true.should == mock_output_stream.write_called
  end
end
