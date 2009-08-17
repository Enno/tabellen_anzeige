require 'spaltenwahl_controller'
class TabelleController < ApplicationController
  set_model 'TabelleModel'
  set_view 'TabelleView'
  set_close_action :exit

  def spaltenwahl_btn_action_performed
    spaltenwahl_controller = SpaltenwahlController.instance
    spaltenwahl_controller.spalten_eintragen :alle => %w[spalte1 sp2 dritte letzte],
                                             :aktive => %w[sp2 letzte]
    p :vor_dialog_open
    spaltenwahl_controller.open
    p :dialog_closed
    spalten = spaltenwahl_controller.aktive_spalten
    p [:vom_dialog_erhalten=, spalten]
    #spaltenwahl_controller.dispose
  end

  def excel_erstellen_btn_action_performed
    eg = ExcelGenerator.new( model.daten_modell )
  end

  def anzeigen_btn_action_performed
    update_model   view_model, :daten_pfad
    p model.daten_pfad
    #model.daten_pfad =
    signal :neuer_daten_pfad
    update_view
  end
end
