
p $:
require 'ffmath'
require 'daten_modell'

class TabelleModel
   attr_reader :daten_pfad, :daten_modell
   
  def initialize
    super
    @daten_modell = nil
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
end
