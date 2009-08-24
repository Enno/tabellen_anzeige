require 'spaltenwahl_controller'
require "#{File.dirname(File.dirname(File.dirname(__FILE__)))}/lib/rcomp/excel_export/lib/export_into_excel"
require 'dateiauswahl_controller'
require 'info_controller'
require 'bestaetigung_controller'

FILE_PATH = Dir.getwd + '/data.xls'

class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view 'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    spaltenwahl_controller = SpaltenwahlController.instance
    spaltenwahl_controller.spalten_eintragen :alle => model.alle_spalten_namen,
      :aktive => model.alle_spalten_namen #model.aktive_spalten_namen(model.selected_columns) #%w[sp2 letzte]
    #TODO: auch die aktiven, aber als speicher model.aktive_spalten_namen
    # bei init auf alle setzen
    p :vor_dialog_open
    spaltenwahl_controller.open
    p :dialog_closed
    aktive_spalten = spaltenwahl_controller.aktive_spalten
    p [:vom_dialog_erhalten=, aktive_spalten]
    inaktive_spalten = model.alle_spalten_namen - aktive_spalten
    aktive_spalten_auswahl(aktive_spalten, inaktive_spalten)

    #spaltenwahl_controller.dispose
    #TODO: nur ausgewaehlte spalten anzeigen (Jtable optionen durchsuchen) mithilfe breite auf 0
  end

  def aktive_spalten_auswahl(aktive_spalten, inaktive_spalten)
    update_model view_model, :aktive_spalten, :inaktive_spalten
    model.aktive_spalten = aktive_spalten
    model.inaktive_spalten = inaktive_spalten
    signal :aktive_spalten_signal
    update_view
  end

  def exportieren_button_action_performed
    eg = ExportIntoExcel.new(FILE_PATH)
    update_model view_model, :daten_modell
    eg.get_data(daten_modell)
    label = "Exportieren erfolgreich (#{FILE_PATH})"
    open_info_dialog(label)
    update_view
  end

  def exportieren_menuitem_action_performed
    exportieren_button_action_performed
  end

  def beenden_menuitem_action_performed
    label = "Programm beenden?"
    close if open_bestaetigung_dialog(label)
  end
  
  def anzeigen_btn_action_performed
    update_model   view_model, :daten_pfad
    #p model.daten_pfad
    signal :neuer_daten_pfad
    update_view
  end

  def exportieren_nach_menuitem_action_performed
    dateiauswahl_controller = DateiauswahlController.instance
    dateiauswahl_controller.open
    destination_path = dateiauswahl_controller.get_destination_path
    eg = ExportIntoExcel.new(destination_path)
    eg.get_data(daten_modell)
    label = "Exportieren erfolgreich (#{destination_path})"
    open_info_dialog(label)
    update_view
  end

  def open_info_dialog(label)
    info_dialog = InfoController.instance
    info_dialog.set_label(label)
    info_dialog.open
  end

  def open_bestaetigung_dialog(label)
    bestaetigung_dialog = BestaetigungController.instance
    bestaetigung_dialog.set_label(label)
    bestaetigung_dialog.open
    dialog_result = bestaetigung_dialog.dialog_result
    return dialog_result
  end

  def daten_modell
    model.daten_modell
  end
end