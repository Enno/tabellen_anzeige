require 'spaltenwahl_controller'
require "#{File.dirname(File.dirname(File.dirname(__FILE__)))}/lib/rcomp/excel_export/lib/export_into_excel"
require 'dateiauswahl_controller'

FILE_PATH = Dir.getwd + '/data.xls'

class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view 'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    spaltenwahl_controller = SpaltenwahlController.instance
    spaltenwahl_controller.spalten_eintragen :alle => model.alle_spalten_namen,
      :aktive => %w[sp2 letzte]
    #TODO: auch die aktiven, aber als speicher model.aktive_spalten_namen
    # bei init auf alle setzen
    p :vor_dialog_open
    spaltenwahl_controller.open
    p :dialog_closed
    spalten = spaltenwahl_controller.aktive_spalten
    p [:vom_dialog_erhalten=, spalten]
    #spaltenwahl_controller.dispose
    #TODO: nur ausgewaehlte spalten anzeigen (Jtable optionen durchsuchen) mithilfe breite auf 0
  end

  def exportieren_button_action_performed
    eg = ExportIntoExcel.new(FILE_PATH)
    update_model view_model, :daten_modell
    eg.exportieren(model.daten_modell)
    #signal :daten_modell_select_all
    update_view
  end

  def anzeigen_btn_action_performed
    update_model   view_model, :daten_pfad
    p model.daten_pfad
    #model.daten_pfad =
    signal :neuer_daten_pfad
    update_view
  end

  def exportieren_nach_menuitem_action_performed
    dateiauswahl_controller = DateiauswahlController.instance
    #TODO: hier currentdir vorbelegen (siehe spaltenwahl)
    dateiauswahl_controller.open
    zielpfad = dateiauswahl_controller.uebergebe_zielpfad
    eg = ExportIntoExcel.new(zielpfad)
    eg.exportieren(daten_modell)
  end

  def daten_modell
    model.daten_modell
  end
end

