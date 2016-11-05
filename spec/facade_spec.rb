require 'spec_helper'

describe "POI.Facade" do
  it "should create specialized methods for boolean methods, getters, and setters" do
    book = POI::Workbook.create('foo.xlsx')
    sheet = book.create_sheet
    sheet.respond_to?(:column_broken?).should be true
    sheet.respond_to?(:column_hidden?).should be true
    sheet.respond_to?(:display_formulas?).should be true
    sheet.respond_to?(:display_gridlines?).should be true
    sheet.respond_to?(:selected?).should be true
    sheet.respond_to?(:column_breaks).should be true
    sheet.respond_to?(:column_break=).should be true
    sheet.respond_to?(:autobreaks?).should be true
    sheet.respond_to?(:autobreaks=).should be true
    sheet.respond_to?(:autobreaks!).should be true
    sheet.respond_to?(:column_broken?).should be true
    sheet.respond_to?(:fit_to_page).should be true
    sheet.respond_to?(:fit_to_page?).should be true
    sheet.respond_to?(:fit_to_page=).should be true
    sheet.respond_to?(:fit_to_page!).should be true

    sheet.respond_to?(:array_formula).should_not be true
    sheet.respond_to?(:array_formula=).should_not be true
    
    row = sheet[0]
    row.respond_to?(:first_cell_num).should be true
    row.respond_to?(:height).should be true
    row.respond_to?(:height=).should be true
    row.respond_to?(:height_in_points).should be true
    row.respond_to?(:height_in_points=).should be true
    row.respond_to?(:zero_height?).should be true
    row.respond_to?(:zero_height!).should be true
    row.respond_to?(:zero_height=).should be true

    cell = row[0]
    cell.respond_to?(:boolean_cell_value).should be true
    cell.respond_to?(:boolean_cell_value?).should be true
    cell.respond_to?(:cached_formula_result_type).should be true
    cell.respond_to?(:cell_error_value=).should be true
    cell.respond_to?(:cell_value=).should be true
    cell.respond_to?(:hyperlink=).should be true
    cell.respond_to?(:cell_value!).should be true
    cell.respond_to?(:remove_cell_comment!).should be true
    cell.respond_to?(:cell_style).should be true
    cell.respond_to?(:cell_style=).should be true
  end
end
