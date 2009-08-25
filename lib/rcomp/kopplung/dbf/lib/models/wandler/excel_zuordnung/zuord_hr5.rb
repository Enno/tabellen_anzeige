def zeit(jahr, monat)
  "#{jahr}|#{monat}"
end


module ExcelZuordnungHR5
  VORLAGE_MINDEST_VERSION = "0-00-000"
  
  FORMELN = {
    :"BonVertragsÜFormel" => proc { <<Ende }
=VertragsÜ(#{ normale_vts.map {|vt|
                "K#{vt.nr_im_vertrag}!BonEinzelKompÜ"
              }.join(';')
            })
Ende
  }
  VD_zuord = {
    :VsNr      => :vsnr,
    :BeiArt    => :beiart,
    :FkAbw     => :fkabw,
#    :VRZins    => proc { promille(vrzins) rescue nil },
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
#    :DivVerwGlobal => proc {
#                    case vk_haupt.divsl.to_s
#                    when "0", "B" then "B"
#                    when "1", "A" then "A"
#                    else               "V"
#                    end
#                  }

  }
  ST_zuord = {
    :StBer     => :stber,
    :ZgPer     => proc { zeit(zgperj, zgperm) }
    # ,                  :geszins  => {"GesamtZins" => PROZENT_Proc}
  }
  VP_zuord = {
    :GsVP1     => :gsvp1,
    :GebVP1   => proc { zeit(gbjvp1, gbmvp1) },
#    :GebJVP1   => :gbjvp1,
#    :GebMVP1   => :gbmvp1,
    :GsVP2     => :gsvp2,
#    :GebVP2   => :gbjvp2,
    :GebVP2   => proc { zeit(gbjvp2, gbmvp2) }
  }
  VK_zuord = {
#    :Bez       => :komp,
    :Komp      => :komp,
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
    :RenZw     => proc { if vk.beitr_befrei_komp? then 1 else vd.renzw end },
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
    :divsl     => :divsl,
#    :ZinDiv_Vorg=> proc { div_vorgabe(:zindiv).transpose },
#    :RisDiv_Vorg=> proc { div_vorgabe(:risdiv).transpose },
#    :GruDiv_Vorg=> proc { div_vorgabe(:grudiv).transpose },
#    :BeiDiv_Vorg=> proc { div_vorgabe(:beidiv).transpose },
    :Rabatt    => proc { vd.kzbrab }
  }
  VT_zuord = {
    :SatzNr    => proc { nr_im_vertrag },  #"'%02d" % 
    :GV        => proc { umsetz_gv( vk.gv, tarif) },
    :Tarif     => proc { umsetz_tarif( vk.gv, tarif) },
    :vtnr      => :vtnr,
    :Status    => :status,
    :Beginn    => proc { zeit(begj, begm) },
    #:BegM      => :begm,
    :VDauer    => proc { zeit(n, nm) },
    #:DauerM    => :nm,
    :BeiDauer  => proc { zeit(m, mm) },
    #:BtrDauerM => :mm,
    :LeiDauer  => proc { zeit(ldj, ldm) },
    #:LeiDauerM => :ldm,
    :x         => :x,
    :y         => :y,
    :TarMod    => :tarmod,
    :np        => :gewbtrg,
    :zp        => :zp,
    :BR        => proc { br }, #if tarif =~ /^FB[LE]/ },
    :btr       => :btr,
    :JRen      => proc { jren if jren > 0 },
    :TSum      => proc { tsum if jren == 0 and tsum > 0 },
    :ESum      => proc { esum if jren == 0 and tsum == 0 and esum > 0 },
    :sterbg    => :sterbg,
    :rbeg      => :rbeg, #proc { zeit(rbegj, rbegm) },

    # Lambda nur bei klasischen Renten (Rxxx) oder MultiflexPP (Fxxx) letzteres sollte aber vielleicht noch eingeschränkt werden
    :Lambda    => proc { vd.vk_haupt.proz/100 if tarif =~ /^(R|F)/ }
  }
  VA_zuord =  {
    :alpha1    => proc { promille(akw1) if not konstante_abskosten?(tarif) },
    :alpha2    => proc { promille(akw2) if not konstante_abskosten?(tarif) },
    #:gv   => :nix,
    #:tarif=> :nix
  }
  RK_zuord =  {
    :ttt       => proc { [vj, vm] },
    :RmDkk     => :rmdkk,
    :NetDrk    => :netdrk,
    :Rkw       => :rkw,
    :AktivW    => :aktivw,
    :RmAk      => :rmak,
    :LaufAZ    => :laufaz,
    :VwkPr     => :vwkpr,
    :StornoAb  => :stornoab,
    :GezDkk    => :gezdkk
#    :dkk       => proc { if rmdkk == 0 then gezdkk else rmdkk end}
  }

end

