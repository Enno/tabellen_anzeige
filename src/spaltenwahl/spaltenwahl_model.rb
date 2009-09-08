class SpaltenwahlModel
  attr_accessor :alle_spalten, :aktive_spalten_namen

  def aktive_spalten_indices
    aktive_spalten_namen.map do |spalten_name|
      alle_spalten.index(spalten_name)
    end
  end

  def aktive_spalten_indices=(array_von_indices)
    @aktive_spalten_namen = array_von_indices.map do |index|
      alle_spalten[index]
    end
  end
end
