
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

  define_signal :name => :aktive_spalten_signal, :handler => :setze_aktive_verstecke_inaktive_spalten

  def setze_aktive_verstecke_inaktive_spalten(model, transfer)
    model.inaktive_spalten.each do |name|
      col_index = blatt.columnModel.getColumnIndex(name)
      blatt.columnModel.getColumn(col_index).setMinWidth(0)
      blatt.columnModel.getColumn(col_index).setMaxWidth(0)
      blatt.columnModel.getColumn(col_index).setWidth(0)
    end
    model.aktive_spalten.each_with_index do |name, index|
      col_index = blatt.columnModel.getColumnIndex(name)
      blatt.columnModel.getColumn(col_index).setMinWidth(10)
      blatt.columnModel.getColumn(col_index).setMaxWidth(10000)
      blatt.columnModel.getColumn(col_index).setPreferredWidth(400)
      blatt.moveColumn(col_index, index)
    end
    blatt.addColumnSelectionInterval(0, model.aktive_spalten.size - 1)
  end

  map :model => :blatt, :view => "blatt", :using => [nil, :default]
  map :model => :col_model, :view => "blatt.column_model", :using => [:default, :default]
  map :model => :daten_modell, :view => "blatt.model"

end
