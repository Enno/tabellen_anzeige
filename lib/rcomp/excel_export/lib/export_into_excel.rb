require 'spreadsheet'

class ExportIntoExcel
  attr_reader :data_model

  def initialize(filepath, data_model)
    Spreadsheet.client_encoding = 'UTF-8'
    @destination = filepath
    @data_model = data_model
    @worksheet_name = 'data'
  end

  def get_all_data
    rowcount_value = data_model.getRowCount
    colcount_value = data_model.getColumnCount
    values = Array.new(rowcount_value) {|i| Array.new(colcount_value, nil)}
    colcount_value.times do |col|
      rowcount_value.times do |row|
        values[row][col] = data_model.getValueAt(row, col) ? check_value_format(data_model, row, col) : nil
      end
    end
    column_indices = (0 ... colcount_value).to_a
    write_into_excel_file(column_indices, values)
  end

  def get_selected_data(active_col_indices)
    rowcount_value = data_model.getRowCount
    values = Array.new(rowcount_value) {|i| Array.new(active_col_indices.size-1, nil)}
    active_col_indices.each_with_index do |col_index, active_col_index|
      rowcount_value.times do |row|
        values[row][active_col_index] = data_model.getValueAt(row, col_index) ? check_value_format(data_model, row, col_index) : nil
      end
    end
    write_into_excel_file(active_col_indices, values)
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
      data_model.getValueAt(row, col)
    end
  end

  def write_into_excel_file(column_indices, values)
    book = Spreadsheet::Workbook.new
    sheet = book.create_worksheet :name => @worksheet_name
    column_indices.each_with_index do |column, column_index|
      name = data_model.getColumnName(column)
      sheet[0, column_index] = name
      sheet.row(0).set_format column_index, Spreadsheet::Format.new(
        :weight => :bold,
        :align  => :center,
        :border => 1,
        :pattern => 1,
        :pattern_fg_color => 'gray')
    end
    values.each_with_index do |row, value_index|
      row.each_with_index do |value, row_index|
        sheet[value_index + 1, row_index] = value
        sheet.row(value_index + 1).set_format row_index, Spreadsheet::Format.new(
          :align => :center,
          :border => 1)
      end
    end
    book.write @destination
  end

end

#TODO: Formeln und Kommentare mit spreadsheet moeglich?
