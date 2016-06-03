# -*- coding: utf-8 -*-
require 'spec_helper'

DOCX_SAMPLE_PATH = TestDataFile.expand_path("simple_with_table.docx")

describe POI::Document do
  it "should open a docx format document and allow access to its content" do
    doc = POI::Document.open(DOCX_SAMPLE_PATH)
    doc.filename.should == DOCX_SAMPLE_PATH
  end
  
  it "should return specific paragraphs" do
    doc = POI::Document.open(DOCX_SAMPLE_PATH)
    doc.paragraphs.size.should == 6
    doc.paragraphs[0].text.should == "This document have one title"
    doc.paragraphs[5].text.should == "This is the final paragraph of content"
    doc.paragraphs[-1].text.should == ""
    doc.paragraphs[10].text.should == ""
  end
  
  it "should return specific tables" do
    doc = POI::Document.open(DOCX_SAMPLE_PATH)
    doc.tables.size.should == 1
    table = doc.tables[0]
    table.rows.size.should == 3
  end
  
  it "can retreive specific table content" do
    doc = POI::Document.open(DOCX_SAMPLE_PATH)
    table = doc.tables[0]
    table.rows[0].cells.size.should == 4
    table.rows[0][0].text.should == 'Name'
    table.rows[0][1].text.should == 'Author'
    table.rows[1][2].text.should == '1997'
    table.rows[2][3].text.should == 'NA'
  end
end