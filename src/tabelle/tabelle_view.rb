require 'daten_modell'

class TabelleView < ApplicationView
  #set_java_class 'TabelleFrame'
  set_java_class 'tabelle.TabelleFrame'
  #set_java_class 'tabelle.TabelleFrame'

  def load


    cmodel = blatt.getColumnModel
    #tbelle.setColumnSelectionAllowed true
    cmodel.setColumnSelectionAllowed true
    blatt.setAutoCreateRowSorter(true)
    cmodel.addColumn javax.swing.table.TableColumn.new

    #blatt.setColumnModel cmodel

    self.blatt.setModel DatenModell.new


  end
end
