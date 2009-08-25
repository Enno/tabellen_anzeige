
#p $:
#require 'ffmath'
require 'daten_modell'

class TabelleModel
   attr_reader :daten_pfad
   attr_accessor :daten_modell
  def initialize
    super
    #@daten_modell = nil
    @daten_modell_dummy = @daten_modell = DatenModellDummy.new
    @spaltenname = []
  end

  def daten_pfad=(daten_pfad)
    p [:dpfad=, daten_pfad]
    begin
      @daten_pfad = daten_pfad.gsub("\\","/")
      daten_pfad = "/dat/GiS/gm/MStar/Ms609/daten/TMFPP" if daten_pfad.empty?
      tabellen_zeilen = TABELLEN_FUNDAMENT.tabelle_fuer_pfad(daten_pfad + ":vk")
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
    daten_modell.getColumnCount.times do |x|
    @spaltenname[x] = daten_modell.getColumnName(x)
    end
  end
#TODO: fehlermeldung beheben
end
