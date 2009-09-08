require 'spreadsheet'

class ExportIntoExcel
  def initialize(filepath)
    Spreadsheet.client_encoding = 'UTF-8'
    @destination = filepath
    @worksheet = 'data'
  end

  def get_all_data(data_model)
    rowcount_value = data_model.getRowCount
    colcount_value = data_model.getColumnCount
    values = Array.new(rowcount_value) {|i| Array.new(colcount_value, nil)}
    column_name = Array.new(colcount_value)
    colcount_value.times do |col|
      column_name[col] = data_model.getColumnName(col).to_s
      rowcount_value.times do |row|
        values[row][col] = data_model.getValueAt(row, col) ? check_value_format(data_model, row, col) : nil
      end
    end
    write_into_excel_file(column_name, values)
  end

  def get_selected_data(data_model, active_columns, active_col_indices)
    rowcount_value = data_model.getRowCount
    colcount_value = data_model.getColumnCount
    values = Array.new(rowcount_value) {|i| Array.new(active_columns.size-1, nil)}
#    column_name = Array.new(colcount_value)
#    colcount_value.times do |col|
#      column_name[col] = data_model.getColumnName(col).to_s
#    end
    active_col_indices.each do |col_index|
      p [:col_index, col_index]
#      col = column_name.index(name)
      rowcount_value.times do |row|
        values[row][col_index] = data_model.getValueAt(row, col_index) ? check_value_format(data_model, row, col_index) : nil
      end
    end
    write_into_excel_file(active_columns, values)
  end

  def check_value_format(data_model, row, col)
    case data_model.getValueAt(row, col)
    when /^[\d]*$/  #integer
      data_model.getValueAt(row, col).to_i
    when /^[\d.,]*$/  #float
      data_model.getValueAt(row, col).to_f
    when /^[\w]*$/  #string (bsp: kommentar)
      data_model.getValueAt(row, col).to_s
    when /^[=A-Z($0-9,)\S]*$/ #formel
      data_model.getValueAt(row, col).to_s
    else
      nil
    end
  end

  def write_into_excel_file(column_name, values)
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet :name => @worksheet
    column_name.each_with_index do |name, y|
      sheet[0, y] = name
      sheet.row(0).set_format y, Spreadsheet::Format.new(
        :weight => :bold,
        :align  => :center,
        :border => 1,
        :pattern => 1,
        :pattern_fg_color => 'gray')
    end
    values.each_with_index do |row, x|
      row.each_with_index do |value, y|
        sheet[x + 1, y] = value
        sheet.row(x + 1).set_format y, Spreadsheet::Format.new(
          :align => :center,
          :border => 1)
      end
    end
    book.write @destination
  end
end

#TODO: Formeln und Kommentare mit spreadsheet moeglich?
