
#vielleicht hilft das, damit die pure-Ruby-Spec-AUsführung die Java-Klasse findet?
#include_class 'tabelle.TabelleFrame'

class TabelleView < ApplicationView
  #set_java_class 'TabelleFrame'
  set_java_class 'tabelle.TabelleFrame'
  
  #nest :sub_view => :spaltenwahl_dialog, :using => [:define_popup_parent, nil]
  def define_popup_parent(view, component, model, transfer)
    p :define_popup_parent
    view.parent_component = @main_view_component
  end

  def load


    cmodel = blatt.getColumnModel
    #tbelle.setColumnSelectionAllowed true
    cmodel.setColumnSelectionAllowed true
    blatt.setAutoCreateRowSorter(true)
    cmodel.addColumn javax.swing.table.TableColumn.new

    #blatt.setColumnModel cmodel

    #self.blatt.setModel DatenModellDummy.new
  end

  map :model => :daten_pfad, :view => "daten_pfad.text"

  define_signal :name => :neuer_daten_pfad, :handler => :neuer_daten_pfad

  def neuer_daten_pfad(model, transfer)
    blatt.model = model.daten_modell
  end

  map :model => :daten_modell, :view => "blatt.model"#, :using => [nil, :default]
  map :model => :alle_spalten, :view => "blatt", :using => [nil, :java_array_to_ruby]
  map :model => :aktive_spalten_indices, :view => "blatt", :using => [nil, :java_array_to_ruby]

  def java_array_to_ruby(java_array_of_indices)
    java_array_of_indices.to_a
  end

  def ruby_array_of_indices_to_java(array_of_indices)
    array_of_indices.to_java(Java::int)
  end

  define_signal :name => :aktive_spalten_signal, :handler => :setze_aktive_spalten

  def setze_aktive_spalten(model, transfer)
    blatt.setColumnSelectionAllowed(true)
    blatt.setRowSelectionAllowed(false)
    blatt.clearSelection()
    aktive_spalten_index = []
    model.aktive_spalten.each do |name|
      aktive_spalten_index << blatt.columnModel.getColumnIndex(name)
    end
    aktive_spalten_index.each do |x|
      blatt.addColumnSelectionInterval(x, x)
    end
  end
end
