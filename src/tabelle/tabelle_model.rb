
#p $:
#require 'ffmath'
require 'daten_modell'

class TabelleModel
  attr_reader :daten_pfad
  attr_accessor :daten_modell, :aktive_spalten, :inaktive_spalten, :alle_spalten, :blatt

  def blatt= jtable
    p jtable
    @blatt = jtable
  end

  def initialize
    super
    #@daten_modell = nil
    @daten_modell_dummy = @daten_modell = DatenModellDummy.new
    @spaltenname = []
    @spaltennamen = []
  end

  def daten_pfad=(daten_pfad)
    p [:dpfad=, daten_pfad]
    begin
      @daten_pfad = daten_pfad
      tabellen_zeilen = TABELLEN_FUNDAMENT.tabelle_fuer_pfad(daten_pfad)
      @daten_modell = DatenModell.new(tabellen_zeilen)
      #p @daten_modell
      p daten_modell.getColumnCount
      p daten_modell.getRowCount
    rescue
      p $!
      puts $!.backtrace
    end
  end

  def alle_spalten_namen
    0.upto(daten_modell.getColumnCount-1) do |x|
      @spaltenname[x] = daten_modell.getColumnName(x).to_s
    end
    return @spaltenname
  end

  def aktive_spalten_array=(active_view)
    begin
      p active_view.inspect, self.aktive_spalten
      active_view.doLayout()#(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
      #blatt.setSelectionModel(ListSelectionModel::MULTIPLE_INTERVAL_SELECTION)
      active_view.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
      active_view.setColumnSelectionAllowed(true)
      active_view.setRowSelectionAllowed(false)
      active_view.clearSelection()

#      daten_modell.inaktive_spalten.each do |name|
#        col_index = active.columnModel.getColumnIndex(name)
#        active_view.columnModel.getColumn(col_index).setMinWidth(0)
#        active_view.columnModel.getColumn(col_index).setMaxWidth(0)
#        active_view.columnModel.getColumn(col_index).setWidth(0)
#      end
#
#      daten_modell.aktive_spalten.each_with_index do |name, index|
#        col_index = active_view.columnModel.getColumnIndex(name)
#        active_view.columnModel.getColumn(col_index).setMinWidth(10)
#        active_view.columnModel.getColumn(col_index).setMaxWidth(10000)
#        active_view.columnModel.getColumn(col_index).setPreferredWidth(400)
#        active_view.moveColumn(col_index, index)
#      end
#      active_view.addColumnSelectionInterval(0, daten_modell.aktive_spalten.size - 1)
      active_view.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
    end
  end

end
