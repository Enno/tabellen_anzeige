
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
    p @col_model
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
end
