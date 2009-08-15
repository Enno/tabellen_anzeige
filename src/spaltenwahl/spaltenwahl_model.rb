class SpaltenwahlModel
  attr_accessor :alle_spalten, :aktive_spalten

  def aktive_spalten_indices
    aktive_spalten.map do |spalten_name|
      alle_spalten.index(spalten_name)
    end
  end

  def aktive_spalten_indices=(array_von_indices)
    @aktive_spalten = array_von_indices.map do |index|
      alle_spalten[index]
    end
  end
end
