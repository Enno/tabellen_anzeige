
#p $:
#require 'ffmath'
require 'daten_modell'

class TabelleModel
  attr_reader :daten_pfad
  attr_accessor :daten_modell, :aktive_spalten, :inaktive_spalten, :alle_spalten
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

  def alle_spalten_namen_indices
    0.upto(daten_modell.getColumnCount-1) do |x|
      @spaltenindex[x] = x
    end
    return @spaltenindex
  end

end
