
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

  map :model => :blatt, :view => "blatt"
  map :model => :daten_modell, :view => "blatt.model"
  map :model => :alle_spalten, :view => "blatt", :using => [nil, :default]

#  define_signal :name => :aktive_spalten_signal, :handler => :setze_aktive_spalten
#
#  def setze_aktive_spalten(model, transfer)
#   model.active
#  end

  map :model => :aktive_spalten_array, :view => "blatt", :using => [nil, :default]
  #map :model => :aktive_spalten, :view => "blatt", :using => [nil, :default]

  #  def setze_aktive_spalten(model, transfer)
  #    blatt.doLayout()#(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
  #    #blatt.setSelectionModel(ListSelectionModel::MULTIPLE_INTERVAL_SELECTION)
  #    blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_OFF)
  #    blatt.setColumnSelectionAllowed(true)
  #    blatt.setRowSelectionAllowed(false)
  #    blatt.clearSelection()
  #    deaktiviere_inaktive_spalten(model)
  #    zeige_aktive_spalten(model)
  #    blatt.setAutoResizeMode(javax.swing.JTable::AUTO_RESIZE_ALL_COLUMNS)
  #  end
  #
  #  def deaktiviere_inaktive_spalten(model)
  #    model.inaktive_spalten.each do |name|
  #      col_index = blatt.columnModel.getColumnIndex(name)
  #      blatt.columnModel.getColumn(col_index).setMinWidth(0)
  #      blatt.columnModel.getColumn(col_index).setMaxWidth(0)
  #      blatt.columnModel.getColumn(col_index).setWidth(0)
  #    end
  #  end
  #
  #  def zeige_aktive_spalten(model)
  #    model.aktive_spalten.each_with_index do |name, index|
  #      col_index = blatt.columnModel.getColumnIndex(name)
  #      blatt.columnModel.getColumn(col_index).setMinWidth(10)
  #      blatt.columnModel.getColumn(col_index).setMaxWidth(10000)
  #      blatt.columnModel.getColumn(col_index).setPreferredWidth(400)
  #      blatt.moveColumn(col_index, index)
  #    end
  #    blatt.addColumnSelectionInterval(0, model.aktive_spalten.size - 1)
  #  end
end
