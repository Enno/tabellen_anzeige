
#require 'schmiedebasis'

unless defined? Protokollierer then
  begin 
    require 'models/system_fkt/datei_aufraeumer' 
  rescue LoadError # für ZenTest
    require File.expand_path('datei_aufraeumer', File.dirname(__FILE__))
  end
	
	class Protokollierer
	  TRACE_STUFEN = [ :temp, :info, :hinweis, :fehler, :essenz]
	
	  attr_reader :datei, :dateiname
	   
	  def initialize(art, ist_subprozess)
	    @art = art.to_s
      @ist_subprozess = ist_subprozess

      return
      oeffne_datei
      
	    if art != :stderr then
	      @aufraeumer = DateiAufraeumer.new(
	        :ordner => LOG_DIRNAME,
          :einzel_muster => trace_datei_muster,
          :paket_muster => /ats-sumtrc_(\d{4})-(\d\d)-(\d\d)_(\d\d)(\d\d)_#{@art}(s(ub?)?|m(a(in?)?)?)\.log$/,
          :einzel_zu_paket_name => proc do |einzelname|
            einzelname.sub("trace_", "sumtrc_").
                       sub(/sumtrc_(?=\d\d-\d\d)/, "sumtrc_#{Time.now.year}-") # falls Jahreszahl fehlt
          end
	      )
      end  	            
      
      @datei.puts "\n\n"
      @datei.puts "=============================================================\n\n"

	  end
    
    def oeffne_datei
	    @datei = if @art=="stderr" then
	      $stderr
	    else 
        @dateiname = baue_dateiname(@art + (@ist_subprozess ? "sub" : "main"))
	      File.new(@dateiname, "a")
	    end
    end      
	  
    def trace_datei_muster
      /ats-trace_(\d\d)-(\d\d)_(\d\d)(\d\d)_#{@art}(s(ub?)?|m(a(in?)?)?)\.log$/
    end
    
    def setze_limits(limit_hash)
      @aufraeumer.setze_konfig(limit_hash)
    end
    
    def baue_dateiname(suffix)
      LOG_DIRNAME + "/ats-trace_"+ Time.now.strftime("%m-%d_%H%M") + "_#{suffix[0,6]}.log"
    end
    
    def xx_tracedatei?(name)
      name =~ /ats-trace_(\d\d)-(\d\d)_(\d\d)(\d\d)_#{@art}(s(ub?)?|m(a(in?)?)?)\.log$/
    end
    private :xx_tracedatei?
    
	  def trace_allgemein(stufe, wo, was=:nix_und_gar_nix, &codeblock)
	    return was if stufe.to_s > $trace_stufe.to_s
	
	    ausgabe = was.inspect
      if codeblock then
	      erg = begin
	        codeblock.call
	      rescue
	        "Fehler im trc-Code: #{$!}\n" +
	        caller.last(7).join("\n")
	      end
        if :nix_und_gar_nix == was
  	      was = erg           
          ausgabe = if was.is_a?(String) then
            was
          else
            was.inspect
          end          
        end
	    end
	
	    aufruf_info = begin
	#=begin
	      aufrufer = caller[1]
	      if aufrufer and aufrufer =~ /(^|[\/\\])([^\/\\]+):(\d+)(: *in +[`'](.+)[''])?/ then
	        zeile = $3.to_i
	        methname = $5
	        datei = $2.sub(/.rb$/,"")
	
	        if methname then
	          begrenzung_methname_size = 22
	          if methname.size > begrenzung_methname_size then
	            wegzunehmen = methname.size - begrenzung_methname_size
	            teile = methname.split("_")
	            pos_teile = []
	            teile.each_with_index{|teil,idx| pos_teile << [idx,teil]}
	            pos_teile = pos_teile.sort_by {|(idx,teil)| teil.size}
	            fertige_pos_teile = []
	            while not pos_teile.empty? do
	              summe_size_uebrig = pos_teile.inject(0){|sum,(pos,teil)| sum+teil.size}
	              (idx,teil) = pos_teile.shift
	              hier_wegzunehmen = (wegzunehmen*teil.size).to_f / summe_size_uebrig
	              hier_wegzunehmen = (hier_wegzunehmen + 0.5).to_i
	              neue_teilsize = teil.size - hier_wegzunehmen
	              #debug trace_puts [teil, wegzunehmen, summe_size_uebrig, hier_wegzunehmen].inspect
	              if neue_teilsize < 2 then
	                neue_teilsize = [2, teil.size].min
	                hier_wegzunehmen = teil.size - neue_teilsize
	              end
	              if hier_wegzunehmen > 0 then
	                teil.replace(teil[0 .. -1-hier_wegzunehmen])
	                teil[-1,1] = "."
	                wegzunehmen -= hier_wegzunehmen
	              end
	              fertige_pos_teile << [idx,teil]
	            end
	
	            methname = fertige_pos_teile.sort.map {|(pos, teil)| teil}.join("_")
	          end
	          methname = "%-#{begrenzung_methname_size}s" % methname
	          while methname.gsub!(/ {8,8}$/,"\t") do end
	        end
	        gekuerzter_dateiname = datei.size>10 ? datei[0..1]+datei[3..7] : datei[0..6]
	        "%-11s"%(gekuerzter_dateiname+":"+("%d"%zeile)) + "\t#{methname}"
	      else
	        aufrufer
	      end
	#=end
	    end
	
	    stufen_zeichen = stufe.to_s[0,1].upcase
	    jetzt = Time.now
	    zeile = jetzt.strftime("%H%M:%S")+".%03d"%(jetzt.usec/1000)+"#{stufen_zeichen} #{aufruf_info}\t# " + wo.to_s
	
	    if :nix_und_gar_nix != was   # Aufpassen, nicht anderserum, wir wissen ja nicht was "was" ist.
	      zeile += " => " + ausgabe
	    else
	      was = wo
	    end
	    self.puts zeile
	    was
	  end
	
	  def puts(text)
	    f = @datei
      begin
  	    f.puts text
	      f.flush
      rescue
        f.close rescue nil
        f = oeffne_datei
        f.puts "\n\n"
        f.puts "====== !!! Fortsetzung !!! ======\n\n"        
        f.puts text
        f.flush
      end
	    nil
	  end

### TODO Achtung  Code zum deaktivieren
    def puts(text)
      $stdout.puts text
    end

    def setze_limits(limit_hash)
    end


  #  def dateien_zusammenfassen
    
    # gibt nil zurück, wenn nichts mehr zu tun ist  
    def aufraeum_schritt
      trc_info :schritt_vorher
      if @aufraeumer then
        anf_zeit = Time.now
        bericht = nil
        begin
          trc_hinweis :schritt_anfang #, aktion
          bericht = @aufraeumer.next_aufraeum_aktion   
        ensure
          trc_essenz "dauer=",  "%3.3f %s" % [Time.now-anf_zeit, bericht]
          erg = if bericht =~ /fertig$/i then
            nil
          else
            method(:aufraeum_schritt)
          end
        end
      else
        erg = nil
      end
      erg
    end
  
	end # class Protokollierer
end

if __FILE__ == $0
  durchlaufe_unittests($0)
end

