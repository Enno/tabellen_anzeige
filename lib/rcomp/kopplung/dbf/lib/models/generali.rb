if __FILE__ == $0 then
require 'schmiedebasis'
require 'models/system_fkt/dbase_zugriff'
end

require 'models/system_fkt/speicherung'

if not defined?(DivRec) then

  class String
    def links(anzahl)
      self[0,anzahl]
    end

    def rechts(anzahl)
      anzahl = size if anzahl > size
      anzahl > 0 ?
        self[-anzahl..-1] :
        ""
    end

  end

  #$trace_stufe = :temp

  class DbfDat::Vd  # ZenTest SKIP
    def vks
      return @sortierte_vks if defined? @sortierte_vks

      beibefr_vks = unsortierte_vks.select {|vk|
        vk.beitr_befrei_komp?
      }
      #trc_temp :beibefr_vks, beibefr_vks
      if beibefr_vks.size > 1 then
        trc_fehler "Zwei beibefr-Komponenten im Vertrag! vsnr=", [vsnr]
      end 

      @sortierte_vks = unsortierte_vks.to_a - beibefr_vks + beibefr_vks # d.h. ans Ende bringen
    end

  end
  
  class DbfDat::St  # ZenTest SKIP
    def zgper_zeit
      Zeit.jm(zgperj, zgperm)
    end
  end
  
  class DbfDat::Vk  # ZenTest SKIP
    def renten_komp?
      vt1.tarif =~ /^(R...|PAR.|EH..)/
    end

    def beitr_befrei_komp?
      trc_info :umsetz_tarif, umsetz_tarif(gv, vt1.tarif)
      umsetz_tarif(gv, vt1.tarif) =~ /^([BCDH]00[1234]|E[BES]10)$/
         # also B001, EB10, EE10, ES10, B003, C003, D003, H003...
    end

    def bonus?
      trc_temp :divsl, divsl
      divsl =~ /^[BE]/
    end

    def div_vorgabe(art)
      begin
        msys = MstarSys.new(vt1.gv_tarif_le)
        satz_nr = case vp.gsvp1
        when "M" then msys.div_nr_mann
        when "F" then msys.div_nr_frau
        end

        dkz = st.divdkz
        divjahr = if dkz and dkz > "" then
          "200"+dkz
        else
          vt1.begj
        end
        dr = MstarDiv.fuer_jahr(divjahr).drec(satz_nr, art, vt1.status)
      rescue
        $ats.konsole.meldung("#{$!}")
        dr = DIVREC_NIL
      end
      [[dr.wartzut, dr.bzpkt, dr.satz]]
    end

  end

  class DbfDat::Vt  # ZenTest SKIP
    def gv_tarif_le
      umsetz_gv(vk.gv, tarif) + umsetz_tarif(vk.gv, tarif) + (vd.beiart>0 ? "L" : "E")
    end

    def beg_zeit
      Zeit.jm(begj, begm)
    end
  end

  def umsetz_tarif(gv, tarif)

    case tarif
    when /R00[1234]/
     then'R001'
    when /R0([1234])\1/
     then'R011'
    when /([TBCDH])Z0([1-9])/
      erg =  "#{$1}00#{$2}"
      if gv !~ /E.0/ then
        erg
      else
        tarif
      end
    when /(TS|[BCDH][KR])Z([1-9])/
      then            "#{$1}0#{$2}"
    else
      tarif
    end
  end

  def umsetz_gv(gv, tarif)
    case gv+tarif
    when /G..U003/
      then 'G00'
    when /([EG])..BU0[123]/
      then "#{$1}00"
    when /([EG])..(BUZ|RIZ|UZV)[123]/
      then "#{$1}00"
    else
      gv
    end
  end

  def konstante_abskosten?(tarif)
    tarif =~ /^(U001|(PR|E[BES])1\d)$/
       # also U001, PR10, EB10, EE10, ES10, EB11, ...
  end

  class DivRec < Struct.new(:zpkbzpkt, :bzpkt, :fbzpkt, :satz, :fpromill, :wartzut, :zpkzut, :dienst)
  end

  DIVREC_NIL = DivRec.new

  def dateiname_egal_ob_gpf_oder_gle(dateiname)
    dn = dateiname.dup
    if File.exist?(dn) then
      return dn
    else
      if dn.sub!(/GENERALI\//i, "GPF/") or
         dn.sub!(/GPF\//i, "GENERALI/") then
        return dn if File.exist?(dn)
      end
      raise "MStar-Datei #{dateiname} fehlt"
    end
  end

  class MstarDiv
    @@divrecs = {}
    def self.fuer_jahr(jahr)
      @@divrecs[jahr] ||= new(jahr)
    end

    def initialize(jahr)
      trc_temp :divordner, KONFIG.opts[:MathstarDivOrdner]
      @datei_inhalt = File.read(dateiname_egal_ob_gpf_oder_gle(
                                KONFIG.opts[:MathstarDivOrdner] + "/#{jahr}"))
      trc_temp :s, @datei_inhalt.size
      @saetze_hash = {}
    end

    def drec(satz_nr, typ, status)
      ms_typ = case typ
      when :beidiv then 2
      when :grudiv then 3
      when :risdiv then 4
      when :zindiv then 6
      else raise "Unbekannter DivTyp (#{typ.inspect}), erlaubt sind: :beidiv, :grudiv, :risdiv, :zindiv."
      end
      div_status = case status.to_i
      when 0 then 3  # Einm-beitr
      when 1 then 1  # btrPfl
      when 2 then 3  # bed. btrFr
      when 3 then 4  # Rente
      when 5 then 4  # GarRente
      when 6 then 2  # verm btrFr
      else raise "Unbekannter status (#{status.inspect})"
      end
      trc_info :nr_typ_status, [satz_nr, ms_typ, div_status]
      if @saetze_hash[satz_nr] then
        satz_inhalt = @saetze_hash[satz_nr]
      else
        pos = @datei_inhalt =~ /^#{"%03d"%satz_nr}$/
        satz_inhalt = []
        if pos then
          @datei_inhalt[pos+4 .. -1].each_line do |zeile|
            break if zeile.size < 8
            satz_inhalt << zeile
          end
        end
        @saetze_hash[satz_nr] = satz_inhalt
      end

      trc_temp :satz_inhalt, satz_inhalt

      zeile = satz_inhalt.find {|zeile| zeile =~ /^\+#{ms_typ} #{div_status}/ }
      trc_temp :zeile, zeile

      return DIVREC_NIL if not zeile

      erg = DivRec.new
      erg.zpkbzpkt = zeile[ 5,2].to_i
      erg.bzpkt    = case zeile[ 7,2].to_i
      when  1 then "BP"
      when  2 then "BR"
      when  3 then "NP"
      when  7 then "RBPalt"
      when  8 then "RTPalt"
      when 11 then "Lst"
      when 12 then "adVHGB"
      when 15 then "BPohneStk"
      else zeile[ 7,2].to_i
      end
      erg.fbzpkt   = zeile[ 9,2].to_i
      erg.satz     = zeile[11,5].to_f/100_000
      erg.fpromill = zeile[16,2].to_i

      erg.wartzut = if zeile[18,1] == 'f' then
        zeile[19,1] # #+# achtung, unvollst�ndig!
      else
        zeile[18,2]
      end.to_i
      erg.zpkzut   = zeile[20,1]
      erg.dienst   = zeile[21,1]
      erg
    end

  end

  class MstarSys < Struct.new(:div_nr_mann, :div_nr_frau, :sdiv_nr)
    def initialize(gv_tarif_le)
      gv = gv_tarif_le[0,3]
      sysdatei_name = File.dirname(KONFIG.opts[:MathstarDivOrdner])+
                            "/SYS#{gv}/#{gv_tarif_le[0,7]}\.sys"
      datei_inhalt = File.read(dateiname_egal_ob_gpf_oder_gle(sysdatei_name))
      trc_temp :datei_inhalt, datei_inhalt.size
      gefunden = datei_inhalt.scan(/^#{gv_tarif_le}\.SYS\n(.+)(\Z|.{8}\.SYS)/m).first
      if gefunden then
        abschnitt_inhalt = gefunden.first
        #trc_temp :abschnitt_inhalt, abschnitt_inhalt
        zweite_zeile = abschnitt_inhalt.split("\n")[1]
        trc_temp :zweite_zeile, zweite_zeile
        self.div_nr_mann = zweite_zeile[54,3].to_i
        self.div_nr_frau = zweite_zeile[57,3].to_i
        self.sdiv_nr = zweite_zeile[72,3].to_i
      else
        raise "Tarifid #{gv_tarif_le} nicht in #{sysdatei_name}"
      end
    end

  end

end

if __FILE__ == $0 then

#require 'fixture_generell'


  class KonsoleProxy
    def meldung(text)
      trc_hinweis :konsolemeldung, text
    end

  end

  class EinfacherTestschmiedeProxy
    @@konsole = KonsoleProxy.new
    def konsole
      @@konsole
    end

  end

  $ats = EinfacherTestschmiedeProxy.new

def generiere_divsatz_nummern
  exl = excel_zugriff.aktive_oder_neue_instanz
  exl.visible = true
  pdmappe = excel_zugriff.oeffne_oder_aktiviere("D:/GiS/gm/Exstar/E2/ExstarPD.xls", :raise)

  pdblatt = pdmappe.Sheets("PrDb")

  id_spalte = pdblatt.Columns(1)

  trc_info :id_spalte, id_spalte.Address

  zeile = 1

  mann_spnr = pdblatt.Range("PD_DivNrM").Column
  frau_spnr = pdblatt.Range("PD_DivNrF").Column

  loop do
      zeile += 1
      break if zeile > 3000

      tarifid = id_spalte.Cells(zeile,1).Value
      break if tarifid == ""
      trc_info :tarifid, tarifid

      msys = MstarSys.new(tarifid) rescue next

      satz_mann = msys.div_nr_mann
      pdblatt.Cells(zeile, mann_spnr).Value = satz_mann
      satz_frau = msys.div_nr_frau
      pdblatt.Cells(zeile, frau_spnr).Value = satz_frau
  end
end

def generiere_sdiv_nummern
  exl = excel_zugriff.aktive_oder_neue_instanz
  exl.visible = true
  pdmappe = excel_zugriff.oeffne_oder_aktiviere("D:/GiS/gm/Exstar/E2/ExstarPD.xls", :raise)

  pdblatt = pdmappe.Sheets("PrDb")

  id_spalte = pdblatt.Columns(1)

  trc_info :id_spalte, id_spalte.Address

  zeile = 1

  spnr = pdblatt.Range("PD_SDivNr").Column

  loop do
      zeile += 1
      break if zeile > 3000

      tarifid = id_spalte.Cells(zeile,1).Value
      break if tarifid == ""
      trc_info :tarifid, tarifid

      msys = MstarSys.new(tarifid) rescue next

      sdiv_nr = msys.sdiv_nr
      pdblatt.Cells(zeile, spnr).Value = sdiv_nr
  end
end

generiere_sdiv_nummern

=begin
module TestDbfDat
  class TestVt #< Test::Unit::TestCase
    def setup
      DbfDat::oeffnen(FIX_EINFACH_DIRNAME + "/MStarDaten1/cGla", "s")
      @vd = DbfDat::Vd.find("GK_715")
      assert @vd
    end

    def test_gv_tarif_le
      vk = @vd.vk_haupt
      vt = vk.vt1
      assert_equal "G82E001L", vt.gv_tarif_le
      assert_equal 4, @vd.vks.size
    end

  end

  class TestVk < Test::Unit::TestCase
    def setup
      DbfDat::oeffnen(FIX_EINFACH_DIRNAME + "/MStarDaten1/cGla", "s")
      @vd = DbfDat::Vd.find("GK_715")
      assert @vd
    end

    def test_div_vorgabe
      vk = @vd.vk_haupt
      assert_equal "G82E001L", vk.vt1.gv_tarif_le
      assert_equal [[nil, nil, nil]], vk.div_vorgabe(:beidiv)
      assert_equal [[nil, nil, nil]], vk.div_vorgabe(:grudiv)
      assert_equal [[nil, nil, nil]], vk.div_vorgabe(:risdiv)
      assert_equal [[2, "adVHGB", 0.0185]], vk.div_vorgabe(:zindiv)
      assert_equal 4, @vd.vks.size
    end

  end
end
=end

end