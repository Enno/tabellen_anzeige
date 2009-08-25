require 'spreadsheet'

class ExportIntoExcel
  def initialize(filepath)
    Spreadsheet.client_encoding = 'UTF-8'
    @file = filepath
    puts "File:" + @file.to_s
    puts @file.inspect
    @worksheet = 'data'
  end

  def exportieren(daten_modell)
    blatt = daten_modell
    rowcount_value = blatt.getRowCount
    colcount_value = blatt.getColumnCount
    wert = Array.new(rowcount_value) {|i| Array.new(colcount_value, nil)}
    column_name = Array.new(colcount_value)
    colcount_value.times do |col|
      column_name[col] = blatt.getColumnName(col)
      #puts column_name[col]
      rowcount_value.times do |row|
        wert[row][col] = blatt.getValueAt(row, col)
      end
    end

    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet :name => @worksheet
    #puts wert.inspect
    column_name.each_with_index do |name, y|
    sheet[0, y] = name
    sheet.row(0).default_format = Spreadsheet::Format.new :weight => :bold, :size => 14, :align => :center, :border => 1
    end
    wert.each_with_index do |zeile, x|
      zeile.each_with_index do |wert, y|
        sheet[x + 1, y] = wert
      end
    end
    book.write @file
  end
end

#TODO: case fuer integer, float etc (datenvorverarbeitung)
#TODO: refactoring
#TODO: formatieren (grau fuer spaltenueberschriften) und nur fuer belegte zellen
#TODO: Formeln und Kommentare mit spreadsheet moeglich?
#TODO: dialog zur auswahlbestaetigung