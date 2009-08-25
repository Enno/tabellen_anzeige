
 

module ExcelZuordnungNG1
  VORLAGE_MINDEST_VERSION = "2-10-000"
  
  FORMELN = {
    :"BonVertragsÜFormel" => proc { <<Ende }
=VertragsÜ(#{ vts.map {|vt|
                "K#{vt.blatt_nr}!BonEinzelKompÜ"
              }.join(';')
            })
Ende
  }
  VD_zuord = {
    :BeiArt    => :beiart,
    :FkAbw     => :fkabw,
    :VRZins    => proc { promille(vrzins) rescue nil },
    #:ratzus   => :nix, #"RatZus",
    :PersAnz   => :prsanz,
    :VertAbsk  => :vertakj,
    :Stueck    => proc {
                    if vk_haupt.gv.links(1) == "E" then
                      20
                    else
                      stueck
                    end
                  },
    :DivVerwGlobal => proc {
                    case vk_haupt.divsl.to_s
                    when "0", "B" then "B"
                    when "1", "A" then "A"
                    else               "V"
                    end
                  }

  }
  ST_zuord = {
    :StBer     => :stber,
    # ,                  :geszins  => {"GesamtZins" => PROZENT_Proc}
  }
  VP_zuord = {
    :GsVP1     => :gsvp1,
    :GebJVP1   => :gbjvp1,
    :GebMVP1   => :gbmvp1,
    :GsVP2     => :gsvp2,
    :GebJVP2   => :gbjvp2,
    :GebMVP2   => :gbmvp2
  }
  VK_zuord = {
    :Bez       => :komp,
    :KolKz     => :kolkz,
    :VerlDauer => :verldj,
    #:proz     => :nix,
    :RisKlVp1  => proc {
                    case rikvp1
                      when nil then nil
                      when Numeric then "" # alte Formate? -- unklar
                      else (rikvp1.links(1)=='1' ? 'A' : 'N') + ':' + rikvp1.rechts(1)
                    end
                  },
    :MZPrm     => proc { promille(mzprm) },
    :SZPrm     => proc { promille(szprm) },
    :Rzw       => proc { if vk.beitr_befrei_komp? then 1 else vd.renzw end },
    :Prozent   => proc {
                    if komp == "A1" then 1.00 else proz/100 end
                  },
    :ga_ea     => proc {
                   trc_temp "self.endalt", self.endalt
                   trc_temp "endalt", endalt
                   if (endalt||0) > 0 then
                     endalt
                   elsif komp == 'A1' then
                     gazeit
                   else
                     gaz_hk = (vd.vk_haupt.gazeit || 0 rescue 42)
                     trc_info :gaz_hk, gaz_hk
                     if gaz_hk > 0 and vk.renten_komp? then
                       gaz_hk
                     else
                       nil
                     end
                   end
                 },
    :ZinDiv_Vorg=> proc { div_vorgabe(:zindiv).transpose },
    :RisDiv_Vorg=> proc { div_vorgabe(:risdiv).transpose },
    :GruDiv_Vorg=> proc { div_vorgabe(:grudiv).transpose },
    :BeiDiv_Vorg=> proc { div_vorgabe(:beidiv).transpose },
    :Rabatt    => proc { vd.kzbrab }
  }
  VT_zuord = {
    :GV        => proc { umsetz_gv( vk.gv, tarif) },
    :Tarif     => proc { umsetz_tarif( vk.gv, tarif) },
    :BegJ      => :begj,
    :BegM      => :begm,
    :DauerJ    => :n,
    :DauerM    => :nm,
    :BtrDauerJ => :m,
    :BtrDauerM => :mm,
    :LeiDauerJ => :ldj,
    :LeiDauerM => :ldm,
    :BeitrRate => proc { br if tarif =~ /^FB[LE]/ },
    :JRen      => proc { jren if jren > 0 and not vk.beitr_befrei_komp? },
    :TSumme    => proc { tsum if jren == 0 and tsum > 0 and not vk.beitr_befrei_komp?},
    :ESumme    => proc { esum if jren == 0 and tsum == 0 and esum > 0 and not vk.beitr_befrei_komp? },

    # Lambda nur bei klasischen Renten (Rxxx) oder MultiflexPP (Fxxx) letzteres sollte aber vielleicht noch eingeschränkt werden
    :Lambda    => proc { vd.vk_haupt.proz/100 if tarif =~ /^(R|F)/ }
  }
  VA_zuord =  {
    :alpha1    => proc { promille(akw1) if not konstante_abskosten?(tarif) },
    :alpha2    => proc { promille(akw2) if not konstante_abskosten?(tarif) },
    #:gv   => :nix,
    #:tarif=> :nix
  }
  RK_zuord = {
    :RmAbsk   => :rmak,
    :LaufAZ   => :laufaz,
    :VwkPr    => :vwkpr,
    :RatzusVwk=> :rzvk,
    :StornoAb => :stornoab
  }

end

