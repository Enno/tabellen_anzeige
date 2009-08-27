
#p $:
#require 'ffmath'
require 'daten_modell'

import javax.swing.table.DefaultTableColumnModel

class TabelleModel
  attr_reader :daten_pfad
  attr_accessor :daten_modell, :col_model, :aktive_spalten, :inaktive_spalten


  def blatt= jtable
    p jtable
    @col_model = jtable.getColumnModel
    p @col_model
    @blatt = jtable
    @blatt.doLayout()
    @blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
    @blatt.setColumnSelectionAllowed(true)
    @blatt.setRowSelectionAllowed(false)
    @blatt.clearSelection()

    @blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
  end

  def initialize
    super
    p :init
    @col_model = DefaultTableColumnModel.new
#    @normales_column_model = false
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

#  def col_model= cm
#    @normales_column_model = true
#    @col_model = cm
#  end
  
  def col_model
    setze_aktive_verstecke_inaktive_spalten # if @normales_column_model
    @col_model
  end

  def inaktive_spalten= spalten_array
    p spalten_array
    @inaktive_spalten = spalten_array
  end

  def setze_aktive_verstecke_inaktive_spalten
    cm = @col_model
    p ["akt/inakt", aktive_spalten, inaktive_spalten]
    return unless aktive_spalten
    inaktive_spalten.each do |name|
      col_index = cm.getColumnIndex(name)
      cm.getColumn(col_index).setMinWidth(0)
      cm.getColumn(col_index).setMaxWidth(0)
      cm.getColumn(col_index).setWidth(0)
    end
    aktive_spalten.each_with_index do |name, index|
      col_index = cm.getColumnIndex(name)
      cm.getColumn(col_index).setMinWidth(10)
      cm.getColumn(col_index).setMaxWidth(10000)
      cm.getColumn(col_index).setPreferredWidth(400)
      #blatt.moveColumn(col_index, index)
    end
    #blatt.addColumnSelectionInterval(0, model.aktive_spalten.size - 1)
  end


end
