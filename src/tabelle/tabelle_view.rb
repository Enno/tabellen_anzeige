
#vielleicht hilft das, damit die pure-Ruby-Spec-AUsführung die Java-Klasse findet?
#include_class 'tabelle.TabelleFrame'

import javax.swing.JTable;

class TabelleView < ApplicationView
  #set_java_class 'TabelleFrame'
  set_java_class 'tabelle.TabelleFrame'
  
  #nest :sub_view => :spaltenwahl_dialog, :using => [:define_popup_parent, nil]
  def define_popup_parent(view, component, model, transfer)
    p :define_popup_parent
    view.parent_component = @main_view_component
  end

  def load
    p :load
    blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
    blatt.setAutoCreateRowSorter(true)
    blatt.setColumnSelectionAllowed(true)
    blatt.setRowSelectionAllowed(false)
    blatt.clearSelection()
    blatt.doLayout()
  end

  map :model => :daten_pfad, :view => "daten_pfad.text"

  define_signal :name => :neuer_daten_pfad, :handler => :neuer_daten_pfad

  def neuer_daten_pfad(model, transfer)
    blatt.model = model.daten_modell
  end

  map :model => :blatt,        :view => "blatt", :using => [nil, :default]
  map :model => :col_model,    :view => "blatt.column_model"
  map :model => :daten_modell, :view => "blatt.model"

end
