# To change this template, choose Tools | Templates
# and open the template in the editor.
require 'spreadsheet'
require 'export_into_excel'
require File.dirname(__FILE__) + '/fixtures/daten_modell_1'


describe ExportIntoExcel do
  filepath = File.dirname(File.dirname(__FILE__)) + "/temp/test_temp/test_data.xls" #.gsub('\\','/')
  daten_modell_dummy = DatenModellDummy.new
  File.delete(filepath) rescue nil
  @export_into_excel = ExportIntoExcel.new(filepath)
  @export_into_excel.get_data(daten_modell_dummy)
  before(:each) do
    
  end

  it "xls datei angelegt?" do
    File.exist?(filepath).should == true
  end

  it "sollte integer auslesen" do
    z, s = 1, 1
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == daten_modell_dummy.getValueAt(z,s).to_i
  end

  it "sollte float auslesen" do
    z, s  = 3, 3
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == daten_modell_dummy.getValueAt(z,s).to_f
  end

  it "sollte strings auslesen" do
    z, s = 2, 3 
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == daten_modell_dummy.getValueAt(z,s).to_s
  end

  it "sollte nil auslesen" do
    z, s = 4, 4
    book = Spreadsheet.open filepath
    book.worksheet(0).row(z+1)[s].should == daten_modell_dummy.getValueAt(z,s)
  end

  it "sollte spaltennamen auslesen" do
    book = Spreadsheet.open filepath
    book.worksheet(0).row(0)[2].should == daten_modell_dummy.getColumnName(2)
  end


end


