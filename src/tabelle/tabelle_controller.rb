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
      :aktive => model.alle_spalten_namen #bei init alle gesetzt
    spaltenwahl_controller.open
    @aktive_spalten = spaltenwahl_controller.aktive_spalten
    @inaktive_spalten = model.alle_spalten_namen - @aktive_spalten
    aktive_spalten_auswahl #(aktive_spalten, inaktive_spalten)

    #spaltenwahl_controller.dispose
  end

  def aktive_spalten_auswahl #(aktive_spalten, inaktive_spalten)
    update_model view_model, :aktive_spalten, :inaktive_spalten
    model.aktive_spalten = @aktive_spalten
    model.inaktive_spalten = @inaktive_spalten
    signal :aktive_spalten_signal
    update_view
  end

  def exportieren_button_action_performed
    eg = ExportIntoExcel.new(FILE_PATH)
    update_model view_model, :daten_modell
    eg.get_all_data(daten_modell)
    label = "Exportieren erfolgreich (#{FILE_PATH})"
    open_info_dialog(label)
    update_view
  end

  def exportieren_menuitem_action_performed
    exportieren_button_action_performed
  end

  def beenden_menuitem_action_performed
    label        = "Programm beenden?"
    button1_text = "Confirm"
    button2_text = "Cancel"
    close if open_bestaetigung_dialog(label, button1_text, button2_text)
  end
  
  def anzeigen_btn_action_performed
    update_model   view_model, :daten_pfad
    p model.daten_pfad
    signal :neuer_daten_pfad
    update_view
  end

  def exportieren_nach_menuitem_action_performed
    label        = "Welche Spalten sollen exportiert werden?"
    button1_text = "Alle"
    button2_text = "Ausgewählte"

    dateiauswahl_controller = DateiauswahlController.instance
    dateiauswahl_controller.open
    destination_path = dateiauswahl_controller.get_destination_path
    eg = ExportIntoExcel.new(destination_path)
    case open_bestaetigung_dialog(label, button1_text, button2_text
#        :label        => "Welche Spalten sollen exportiert werden?",
#        :button1_text => "Alle",
#        :button2_text => "Ausgewählte"
      )
    when true
      eg.get_all_data(daten_modell)
      label = "Exportieren erfolgreich (#{destination_path})"
      open_info_dialog(label)
    when false
      if @aktive_spalten
        eg.get_selected_data(daten_modell, @aktive_spalten)
        label = "Exportieren erfolgreich (#{destination_path})"
        open_info_dialog(label)
      else
        label = "Sie haben keine Spalte(n) ausgewählt"
        open_info_dialog(label)
      end
    end
    update_view
  end

  def open_info_dialog(label)
    info_dialog = InfoController.instance
    info_dialog.set_label(label)
    info_dialog.open
  end

  def open_bestaetigung_dialog(label, button1_text, button2_text)
    bestaetigung_dialog = BestaetigungController.instance
    bestaetigung_dialog.set_label(label, button1_text, button2_text)
    bestaetigung_dialog.open
    dialog_result = bestaetigung_dialog.dialog_result
    return dialog_result #true = button1, false = button2
  end

  def daten_modell
    model.daten_modell
  end
end