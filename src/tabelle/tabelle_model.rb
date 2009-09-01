
#p $:
#require 'ffmath'
require 'daten_modell'

import javax.swing.table.DefaultTableColumnModel

class TabelleModel
  attr_reader :daten_pfad
  attr_accessor :daten_modell, :col_model, :aktive_spalten, :inaktive_spalten, :blatt


  def blatt= jtable
    p [:blatt, jtable]
    @col_model = jtable.getColumnModel
    p @col_model
    @blatt = jtable
#    @blatt.doLayout()
    @blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
    @blatt.setColumnSelectionAllowed(true)
    @blatt.setRowSelectionAllowed(false)
    @blatt.clearSelection()
    #@blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
    @blatt
  end

  def initialize
    super
    p :init
    @col_model           = DefaultTableColumnModel.new
    @daten_modell_dummy  = @daten_modell = DatenModellDummy.new
    @spaltenname         = []
    @col_total_width_old = []
    @col_pref_width_old  = []
    @col_min_width_old   = []
    @col_max_width_old   = []
#    @col_width_old       = []
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

  def col_model
    setze_aktive_verstecke_inaktive_spalten 
    @col_model
  end

  def alle_spalten_breiten(cm)
    alle_spalten_namen.each do |name|
      col_index = cm.getColumnIndex(name)
      @col_min_width_old[col_index]   = cm.getColumn(col_index).getMinWidth()       == 0 ? @col_min_width_old[col_index]  : cm.getColumn(col_index).getMinWidth()
      @col_max_width_old[col_index]   = cm.getColumn(col_index).getMaxWidth()       == 0 ? @col_max_width_old[col_index]  : cm.getColumn(col_index).getMaxWidth()
      @col_pref_width_old[col_index]  = cm.getColumn(col_index).getPreferredWidth() == 0 ? @col_pref_width_old[col_index] : cm.getColumn(col_index).getPreferredWidth()
    end
  end
  
  def setze_aktive_verstecke_inaktive_spalten
    jt = @blatt
    cm = @col_model
    p ["akt/inakt", aktive_spalten, inaktive_spalten]
    return unless aktive_spalten
    alle_spalten_breiten(cm)
    inaktive_spalten.each do |name|
      col_index            = cm.getColumnIndex(name)
      cm.getColumn(col_index).setMinWidth(0)
      cm.getColumn(col_index).setMaxWidth(0)
      cm.getColumn(col_index).setWidth(0)
    end
    aktive_spalten.each_with_index do |name, index|
      col_index = cm.getColumnIndex(name)
      cm.getColumn(col_index).setMinWidth(@col_min_width_old[col_index])
      cm.getColumn(col_index).setMaxWidth(@col_max_width_old[col_index])
      cm.getColumn(col_index).setPreferredWidth(@col_pref_width_old[col_index])
      jt.moveColumn(col_index, index)
    end
    jt.addColumnSelectionInterval(0, aktive_spalten.size - 1)
  end


end
