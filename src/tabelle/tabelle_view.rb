
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

    #self.blatt.setModel DatenModell.new


  end

  define_signal :name => :neuer_daten_pfad, :handler => :neuer_daten_pfad

  def neuer_daten_pfad(model, transfer)
    blatt.setModel model.daten_modell
  end

end
