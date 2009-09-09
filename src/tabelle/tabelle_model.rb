
#p $:
#require 'ffmath'
require 'daten_modell'

import javax.swing.table.DefaultTableColumnModel

class TabelleModel
  attr_reader :daten_pfad
  attr_accessor :daten_modell, :col_model, :aktive_spalten_namen, :blatt

  def initialize
    super
    p :init
    @col_model           = DefaultTableColumnModel.new
    @daten_modell_dummy  = @daten_modell = DatenModellDummy.new
    @col_width           = Array.new(@col_model.column_count)
  end


  def blatt= jtable
    #p [:blatt, jtable]
    @col_model = jtable.getColumnModel
    p [:blatt_aktive=, aktive_spalten_namen]
    @blatt = jtable
    @blatt.doLayout()
    @blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
    @blatt.setColumnSelectionAllowed(true)
    @blatt.setRowSelectionAllowed(false)
    @blatt.clearSelection()
    #@blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
    @blatt
  end

  def col_model= cm
    @col_model = cm
  end

  def daten_pfad=(d_pfad)
    p [:dpfad=, d_pfad]
    begin
      @daten_pfad = d_pfad
      @daten_pfad = "/dat/GiS/gm/MStar/MsNeu/DATEN/<AT7V2>" if @daten_pfad.empty?
      @daten_pfad += ":vk" unless @daten_pfad =~ /:\w+$/
      tabellen_zeilen = TABELLEN_FUNDAMENT.tabelle_fuer_pfad(@daten_pfad)
      @daten_modell = DatenModell.new(tabellen_zeilen)
      #p @daten_modell
#      p daten_modell.getColumnCount
#      p daten_modell.getRowCount
    rescue
      p $!
      puts $!.backtrace
    end
  end

  def alle_spalten_namen
    alle_spalten_indices.map do |x|
      daten_modell.getColumnName(x).to_s
    end
  end

  def alle_spalten_indices
    (0 .. daten_modell.getColumnCount-1).to_a
  end

  def aktive_spalten_indices
    aktive_spalten_namen.map {|name| @col_model.getColumnIndex(name) }
  end

  def inaktive_spalten_indices
    alle_spalten_indices - aktive_spalten_indices
  end

  def aktive_spalte?(col_index)
    aktive_spalten_indices.include? col_index
  end


  def speichere_spalten_breiten
    cm = @col_model
    alle_spalten_indices.each do |col_index|
      if cm.column(col_index).preferred_width > 0
        @col_width[col_index] = {
          :min  => cm.column(col_index).min_width,
          :max  => cm.getColumn(col_index).max_width,
          :pref => cm.getColumn(col_index).preferred_width
        }
      end
    end
    p @col_width
    return @col_width
  end
  
  def setze_aktive_verstecke_inaktive_spalten_fuer_view
    p ["akt/inakt", aktive_spalten_namen, alle_spalten_namen - aktive_spalten_namen ]
    return unless aktive_spalten_namen    
    speichere_spalten_breiten
    alle_spalten_indices.each  do |col_index|
      setze_spalten_breite(col_index)
    end
  end

  def setze_spalten_breite(col_index, breite = (aktive_spalte?(col_index) ? :normal : 0))
    breiten_hash = case breite
    when :normal
      @col_width[col_index]
    when 0
      {:min => 0, :max => 0, :pref => 0}
    else
      raise "Breite #{breite.inspect} nicht erlaubt."
    end
    cm = @col_model
    unless breiten_hash.nil?
      cm.getColumn(col_index).min_width      = breiten_hash[:min]
      cm.getColumn(col_index).max_width      = breiten_hash[:max]
      cm.getColumn(col_index).preferred_width = breiten_hash[:pref]
    end
  end


end
