class TabelleModel
   attr_reader :daten_pfad, :daten_modell
   
  def initialize
    super
    @daten_modell = nil
  end

  def daten_pfad=(daten_pfad)
    @daten_pfad = daten_pfad
    @daten_modell = DatenModell.new(daten_pfad)
  end
end
