require 'java'

class SpaltenwahlView < ApplicationView
  set_java_class "spaltenwahl.SpaltenwahlDialog"

  def create_main_view_component
    SpaltenwahlDialog.new(nil, true)
  end

  def load

  end

  map :view => "spaltenliste.list_data", :model => :alle_spalten, :using => [:setze_alle, nil]
  map :view => "spaltenliste",           :model => :alle_spalten, :using => [nil, :hole_alle]

  def setze_alle(array_of_strings)
    array_of_strings.to_java(:String)
  end

  def hole_alle(jlist_object)
    list_model = jlist_object.model
    (0...list_model.size).map do |index|
      list_model.element_at(index)
    end
  end

  map :view => "spaltenliste.selected_indices", :model => :aktive_spalten_indices, :using => [:setze_aktive, :hole_aktive]
  #map :view => "spaltenliste.selected_values", :model => :aktive_spalten_namen, :using => [:setze_aktive, :hole_aktive]

  def setze_aktive(array_of_indices)
    array_of_indices.to_java(Java::int)
  end

  def hole_aktive(java_array_of_indices)
    java_array_of_indices.to_a
  end
end
