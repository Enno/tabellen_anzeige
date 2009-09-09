require 'spreadsheet'
require 'export_into_excel'
require File.dirname(__FILE__) + '/fixtures/daten_modell_1'

filepath = File.dirname(File.dirname(__FILE__)) + "/temp/test_temp/test_data.xls" #.gsub('\\','/')
data_model_dummy = DatenModellDummy.new

describe ExportIntoExcel, "with all columns" do
  before(:all) do
    File.delete(filepath) rescue nil
    @export_into_excel = ExportIntoExcel.new(filepath, data_model_dummy)
    @export_into_excel.get_all_data()
  end


  it "xls datei angelegt?" do
    File.exist?(filepath).should == true
  end

  it "sollte integer auslesen" do
    z, s = 0, 0
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s).to_i
  end

  it "sollte float auslesen" do
    z, s  = 3, 3
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s).to_f
  end

  it "sollte strings auslesen" do
    z, s = 2, 3 
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s).to_s
  end

  it "sollte nil auslesen" do
    z, s = 4, 4
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s)
  end

  it "sollte spaltennamen auslesen" do
    book = Spreadsheet.open filepath
    book.worksheet(0).row(0)[2].should == data_model_dummy.getColumnName(2)
  end

  it "sollte Formeln schreiben" do
    z, s = 3, 2
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s)
  end

  it "sollte Kommentare schreiben" do
    z, s = 4, 0
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == data_model_dummy.getValueAt(z,s).to_s
  end

end


describe ExportIntoExcel, "with selected Columns" do
  before(:all) do
    File.delete(filepath) rescue nil
    @export_into_excel = ExportIntoExcel.new(filepath, data_model_dummy)
  end

  it "row 0 of all columns" do
    active_col_indices = [0, 1, 2 ,3]
    @export_into_excel.get_selected_data(active_col_indices)
    z = 0
    book = Spreadsheet.open filepath
    active_col_indices.each do |s|
      book.worksheet(0).row(z+1)[s].should == data_model_dummy.daten[z][s]
    end
  end

  it "row 0 of columns 1 and 3" do
    active_col_indices = [1, 3]
    @export_into_excel.get_selected_data(active_col_indices)
    z = 0
    book = Spreadsheet.open filepath
    active_col_indices.each_with_index do |s, active_col_indices_index|
      book.worksheet(0).row(z+1)[active_col_indices_index].should == data_model_dummy.daten[z][s]
    end
  end

end


