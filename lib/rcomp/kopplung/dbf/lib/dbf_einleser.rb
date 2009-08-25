
# encoding: utf-8

__DIR__ = File.dirname(__FILE__)

require __DIR__ + "/schmiedebasis"
#require __DIR__ + '/
require 'models/system_fkt/dbase_zugriff'

module EingabeKopplung # eigentlich DbfKopplung --> Plugin-Modell einführen
  module ClassMethods
    def eingabe_dbf(tabellen_art)
      #name = dienst_objekt.class.name.methodize
      _eingabe_dienste << DbfEinleser.new(schluessel, tabellen_art)
      registriere_tabellen_art(tabellen_art, self)
    end

    def alle_zeilen
      _eingabe_dienste.inject([]) do |erg_bisher, dienst|
        alle_zeilen_von_dienst = dienst.alle_schluessel.map do |schl|
          new(schl)
        end
        p [erg_bisher, alle_zeilen_von_dienst]
        erg_bisher + alle_zeilen_von_dienst
      end
    end
  end
end

class DbfEinleser
  attr_reader :tabellen_art
  def initialize(schl_spec, tabellen_art)
    @schluessel_namen = schl_spec
    @tabellen_art = tabellen_art
    @dbf_zeilen = nil
  end



  def lese(*schl)
    daten_holen_falls_noetig
    @dbf_zeilen[schl]
  end

  def alle_schluessel
    daten_holen_falls_noetig
    @dbf_zeilen.keys
  end

#private

  def daten_holen_falls_noetig
    if @dbf_zeilen.nil?
      DbfDat::oeffnen($db_pfad, :s )
      db_zeilen = DbfDat::const_get(@tabellen_art.to_s.capitalize).find(:all)

      @dbf_zeilen = {}
      db_zeilen.each do |db_zeile|
        schl_werte = @schluessel_namen.map {|schl_n|  db_zeile.send(schl_n)}
        @dbf_zeilen[schl_werte] = db_zeile
      end
    end
  end
end