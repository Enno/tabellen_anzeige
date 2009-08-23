require 'spreadsheet'

class ExportIntoExcel
  def initialize(filepath)
    Spreadsheet.client_encoding = 'UTF-8'
    @destination = filepath
    @worksheet = 'data'
  end

  def get_data(data_model)
    @rowcount_value = data_model.getRowCount
    @colcount_value = data_model.getColumnCount
    @value = Array.new(@rowcount_value) {|i| Array.new(@colcount_value, nil)}
    @column_name = Array.new(@colcount_value)
    @colcount_value.times do |col|
      @column_name[col] = data_model.getColumnName(col)
      @rowcount_value.times do |row|
        @value[row][col] = if data_model.getValueAt(row, col)
          check_value_format(data_model, row, col)
        else
          nil
        end
      end
    end
    write_into_excel_file
  end

  def check_value_format(data_model, row, col)
    @value[row][col] = if data_model.getValueAt(row, col) =~ /[0-9.,][^a-zA-Z]/
      if data_model.getValueAt(row, col).match(/[.,]/)
        data_model.getValueAt(row, col).to_f
      else
        data_model.getValueAt(row, col).to_i

      end
    else
          puts data_model.getValueAt(row, col)
          data_model.getValueAt(row, col).to_s
    end
  end

  def write_into_excel_file
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet :name => @worksheet
    #puts wert.inspect
    @column_name.each_with_index do |name, y|
      sheet[0, y] = name
      sheet.row(0).set_format y, Spreadsheet::Format.new(:weight => :bold, :align => :center, :border => 1, :pattern_bg_color => 'gray')
    end
    @value.each_with_index do |row, x|
      row.each_with_index do |value, y|
        sheet[x + 1, y] = value
        sheet.row(x + 1).set_format y, Spreadsheet::Format.new(:align => :center, :border => 1, :pattern_bg_color => 'gray')
      end
    end
    book.write @destination
  end
end

#TODO: formatieren (grau fuer spaltenueberschriften) und nur fuer belegte zellen
#TODO: Formeln und Kommentare mit spreadsheet moeglich?
#TODO: dialog zur auswahlbestaetigung