
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

  def java_array_to_ruby(java_array_of_indices)
    java_array_of_indices.to_a
  end

  def ruby_array_of_indices_to_java(array_of_indices)
    array_of_indices.to_java(Java::int)
  end

  define_signal :name => :aktive_spalten_signal, :handler => :setze_aktive_spalten

  def setze_aktive_spalten(model, transfer)
    blatt.doLayout()#(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
    #blatt.setSelectionModel(ListSelectionModel::MULTIPLE_INTERVAL_SELECTION)
    blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
    blatt.setColumnSelectionAllowed(true)
    blatt.setRowSelectionAllowed(false)
    blatt.clearSelection()
    deaktiviere_inaktive_spalten(model)
    zeige_aktive_spalten(model)
    blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
  end

  def deaktiviere_inaktive_spalten(model)
    inaktive_spalten_index = []
    model.inaktive_spalten.each do |name|
      inaktive_spalten_index << blatt.columnModel.getColumnIndex(name)
    end
    inaktive_spalten_index.each do |x|
      blatt.columnModel.getColumn(x).setMinWidth(0)
      blatt.columnModel.getColumn(x).setMaxWidth(0)
      blatt.columnModel.getColumn(x).setWidth(0)
    end
  end

  def zeige_aktive_spalten(model)
    aktive_spalten_index = []
    model.aktive_spalten.each do |name|
      aktive_spalten_index << blatt.columnModel.getColumnIndex(name)
    end
    aktive_spalten_index.each do |x|
      blatt.addColumnSelectionInterval(x, x)
      blatt.columnModel.getColumn(x).setMinWidth(10)
      blatt.columnModel.getColumn(x).setMaxWidth(2000)
      blatt.columnModel.getColumn(x).setPreferredWidth(110)
    end
  end
end
