require 'spaltenwahl_controller'
#require "#{File.dirname(File.dirname(File.dirname(__FILE__)))}/lib/rcomp/excel_export/lib/export_into_excel"
require "export_into_excel"
require 'dateiauswahl_controller'
require 'info_controller'
require 'bestaetigung_controller'

FILE_PATH = Dir.getwd + '/data.xls'

class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view  'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    update_model view_model, :blatt, :col_model
    spaltenwahl_controller = SpaltenwahlController.instance
    spaltenwahl_controller.spalten_eintragen \
      :alle  => model.alle_spalten_namen,
      :aktive => model.aktive_spalten_namen || model.alle_spalten_namen # beim ersten Mal alle gesetzt
    spaltenwahl_controller.open
    model.aktive_spalten_namen = spaltenwahl_controller.aktive_spalten_namen
    
    model.setze_aktive_verstecke_inaktive_spalten_fuer_view
    update_view
  end

  def exportieren_button_action_performed
    eg = ExportIntoExcel.new(FILE_PATH)
    update_model view_model, :daten_modell
    eg.get_all_data(daten_modell)
    open_info_dialog(
      :label        => "Exportieren erfolgreich (#{FILE_PATH})",
      :button1_text => "Ok"
    )
    update_view
  end

  def exportieren_menuitem_action_performed
    exportieren_button_action_performed
  end

  def beenden_menuitem_action_performed
    close if open_bestaetigung_dialog(
      :label        => "Programm beenden?",
      :button1_text => "Bestätige",
      :button2_text => "Abbruch"
    )
  end
  
  def anzeigen_btn_action_performed
    update_model   view_model, :daten_pfad
    p model.daten_pfad
    signal :neuer_daten_pfad
    update_view
  end

  def exportieren_nach_menuitem_action_performed
    dateiauswahl_controller = DateiauswahlController.instance
    dateiauswahl_controller.open
    #return if dateiauswahl_controller.user_aborted
    destination_path = dateiauswahl_controller.get_destination_path
    eg = ExportIntoExcel.new(destination_path, daten_modell)
    case open_bestaetigung_dialog(
        :label        => "Welche Spalten sollen exportiert werden?",
        :button1_text => "Alle",
        :button2_text => "Ausgewählte"
      )
    when true
      eg.get_all_data()
      open_info_dialog(
        :label        => "Exportieren erfolgreich (#{destination_path})",
        :button1_text => "Ok"
      )
    when false
      if model.aktive_spalten_namen
        active_col_indices = col_model_indices
        p [:col_indices, active_col_indices]
        eg.get_selected_data(active_col_indices)
        open_info_dialog(
          :label      => "Exportieren erfolgreich (#{destination_path})",
          :button1_text => "Ok"
        )
      else
        open_info_dialog(
          :label        => "Sie haben keine Spalte(n) ausgewählt",
          :button1_text => "Ok"
        )
      end
    end
    update_view
  end

  def open_info_dialog(dialog_text_hash)
    info_dialog = InfoController.instance
    info_dialog.set_label(dialog_text_hash)
    info_dialog.open
  end

  def open_bestaetigung_dialog(dialog_text_hash)
    bestaetigung_dialog = BestaetigungController.instance
    bestaetigung_dialog.set_label(dialog_text_hash)
    bestaetigung_dialog.open
    dialog_result = bestaetigung_dialog.dialog_result
    return dialog_result #true = button1, false = button2
  end

  def daten_modell
    model.daten_modell
  end

  def col_model_indices
    col_index = []
    model.aktive_spalten_namen.each do |name|
      col_index << model.col_model.getColumnIndex(name)
    end
    return col_index
  end

end