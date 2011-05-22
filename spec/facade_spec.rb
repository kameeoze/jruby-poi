describe "POI.Facade" do
  it "should create specialized methods for boolean methods, getters, and setters" do
    book = POI::Workbook.create('foo.xlsx')
    sheet = book.create_sheet
    sheet.respond_to?(:column_broken?).should be_true
    sheet.respond_to?(:column_hidden?).should be_true
    sheet.respond_to?(:display_formulas?).should be_true
    sheet.respond_to?(:display_gridlines?).should be_true
    sheet.respond_to?(:selected?).should be_true
    sheet.respond_to?(:column_breaks).should be_true
    sheet.respond_to?(:column_break=).should be_true
    sheet.respond_to?(:autobreaks?).should be_true
    sheet.respond_to?(:autobreaks=).should be_true
    sheet.respond_to?(:autobreaks!).should be_true
    sheet.respond_to?(:column_broken?).should be_true
    sheet.respond_to?(:fit_to_page).should be_true
    sheet.respond_to?(:fit_to_page?).should be_true
    sheet.respond_to?(:fit_to_page=).should be_true
    sheet.respond_to?(:fit_to_page!).should be_true

    sheet.respond_to?(:array_formula).should_not be_true
    sheet.respond_to?(:array_formula=).should_not be_true
    
    row = sheet[0]
    row.respond_to?(:first_cell_num).should be_true
    row.respond_to?(:height).should be_true
    row.respond_to?(:height=).should be_true
    row.respond_to?(:height_in_points).should be_true
    row.respond_to?(:height_in_points=).should be_true
    row.respond_to?(:zero_height?).should be_true
    row.respond_to?(:zero_height!).should be_true
    row.respond_to?(:zero_height=).should be_true

    cell = row[0]
    cell.respond_to?(:boolean_cell_value).should be_true
    cell.respond_to?(:boolean_cell_value?).should be_true
    cell.respond_to?(:cached_formula_result_type).should be_true
    cell.respond_to?(:cell_error_value=).should be_true
    cell.respond_to?(:cell_value=).should be_true
    cell.respond_to?(:hyperlink=).should be_true
    cell.respond_to?(:cell_value!).should be_true
    cell.respond_to?(:remove_cell_comment!).should be_true
    cell.respond_to?(:cell_style).should be_true
    cell.respond_to?(:cell_style=).should be_true
  end
end
