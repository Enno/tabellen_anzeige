# encoding: iso-8859-15

if not defined? DbfDat::VObjektBasis then

require 'schmiedebasis'

gem 'dbf', "= 1.0.5"
require 'dbf'
module DBF # ZenTest SKIP

  class Column # ZenTest SKIP
    alias old_init initialize 
    undef old_init
    def initialize(name, type, length, decimal)
      raise ColumnLengthError, "field length must be greater than 0" unless length > 0
      @name = strip_non_ascii_chars(name).downcase
      @type, @length, @decimal = type, length, decimal
    end
  end
  
  class Table # ZenTest SKIP
    alias old_init initialize 
    undef old_init
    def initialize(filename, options = {})
      @options = {:in_memory => true, :accessors => true}.merge(options)
      
      @record_klasse = @options[:record_klasse] || Record
      #raise "#{@record_klasse.inspect} erbt nicht von DBF::Record" unless @record_klasse.
      @in_memory = @options[:in_memory]
      @accessors = @options[:accessors]
      @data = File.open(filename, 'rb')
      @memo = open_memo(filename)
      reload!
    end
    
    undef get_record_from_file
    def get_record_from_file(index)
      seek_to_record(@db_index[index])
      deleted_record? ? nil : @record_klasse.new(self)
    end
    
    undef get_all_records_from_file
    def get_all_records_from_file
      all_records = []
      0.upto(@record_count - 1) do |n|
        seek_to_record(n)
        all_records << @record_klasse.new(self) unless deleted_record?
      end
      all_records
    end    
  
  end
  
  class Record # ZenTest SKIP

    def unpack_column(column)
      #@data.read(column.length).unpack("a#{column.length}")
      @data.read(column.length).unpack("a#{column.length}").first
    end

    undef define_accessors
    def define_accessors
# @@accessors_defined wird nicht mehr verwendet (Probleme mit Subklassen)
#      return if @@accessors_defined
      return if self.class.accessors_defined
      defined = false  
      @table.columns.each do |column|
        underscored_column_name = underscore(column.name)
        if @table.options[:accessors] && !respond_to?(underscored_column_name)
          self.class.send :define_method, underscored_column_name do
            @attributes[column.name]
          end
          defined = true
#          @@accessors_defined = true
        end
      end
      if defined then
        class << self
          undef accessors_defined if method_defined?(:accessors_defined)
          def self.accessors_defined; true; end
        end
      end
    end
    
    class << self
      undef klassenreset if method_defined? :klassenreset
      def klassenreset
        class << self
          undef accessors_defined if method_defined?(:accessors_defined)
          def accessors_defined
            false
          end
        end
      end
    end
    
    klassenreset
  end
    
end

module ActiveRecord 
  class ConnectionNotEstablished < Exception
  end
  class StatementInvalid < Exception
  end
end

module DbfDat
  trc_caller :class_DbfDat
  EIGENES_AR_CACHING = false
  def self.oeffnen(dateiname, bestand_oder_soll)
    trc_hinweis :beginne_oeffnen_dateiname=, dateiname
    profil_setzen(File.basename(dateiname), bestand_oder_soll)
    setze_dbf_pfad(File.dirname(dateiname))
    trc_info :DbfDat_oeffnen_ok

  end

  def self.schliessen
    VO_KLASSEN.each do |vklasse|
      vklasse.klassenreset
    end
    #VObjektBasis.connection.disconnect!
    #VObjektBasis.remove_connection
  end

  $aktueller_dbf_ordner = nil
  $abfragen_seit_letztem_connect = 0

private
  # returns nil if dbf_pfad unchanged
  def self.setze_dbf_pfad( dbf_ordner )
    if $aktueller_dbf_ordner != dbf_ordner then
      $aktueller_dbf_ordner = dbf_ordner
    end    
  end 
  
  # Diese Prozedur soll sowohl für den Fall funktioniern, dass
  # profil die ".dbf"-Extension trägt, als auch nicht,
  # weiterhin auch für den Fall, dass es ein reines Profil ist
  # oder die Tabellenart und S/B vorne dranhängt.
  # #*# Die letztere Unterscheidung ist im Fall Länge=5 und Länge =4
  # heikel: "vdbes" könnte das Profil "es" oder "vdbes" meinen.
  # Hier wurde die Entscheidung für die erste Variante getroffen,
  # d.h. im Zweifelsfall _wird_ _der_ _Tabellenpräfix_ _erwartet_.
  # #*# Globales Refactoring könnte diese Frage anders lösen!!!
  def self.profil_setzen(profil, bestand_oder_soll)
    trc_hinweis :setprof_profiluebergeben, profil
    if profil =~ /^<(.+)>$/ then
      profil = $1
    else
      profil.sub!(/^(st|gz|rk|zs|v[dpktvbfa])[BS]/i, '') if profil.size > 3
      profil.sub!(/\.dbf$/i, '')
    end
    trc_info :setprof_profilextrahiert, profil

    b_oder_s = bestand_oder_soll.to_s.upcase
    DbfDat::VO_KLASSEN.each { |vklasse|
      prefix = vklasse.name.split("::").last.upcase
      #prefix = vklasse.name.split("::").last.downcase
      vklasse.set_table_name "" + prefix + b_oder_s + profil #+ ".DBF'"
      #vklasse.set_table_name "" + prefix + b_oder_s + profil #+ ".dbf'"
    }
  end
  
  def self.schnell_lesen(tabellen_name, *felder)
    t = DBF::Table.new(File.join($aktueller_dbf_ordner, tabellen_name + ".dbf"))
    records = t.find(:all)
    if not felder.empty?
      records.map do |record|
        felder.inject({}) do |erg_record, feldsym| 
          erg_record[feldsym] = record.attributes[feldsym.to_s.downcase]
          erg_record
        end
      end
    else
      records.map do |record|
        record.attributes
      end
    end
  end
  
  def self.schnell_anzahl(tabellen_name)
    tabellen_name += ".dbf" unless tabellen_name =~ /\.dbf$/i
    dateiname = if tabellen_name =~ /^(\w:)?\// then
      tabellen_name 
    else
      File.join($aktueller_dbf_ordner, tabellen_name)
    end
    t = DBF::Table.new(dateiname)
    t.record_count
  end
  
  class KeySystem
    
    def initialize(key_struktur)
      @struktur = key_struktur
      @pfade = {}
      notiere_pfade(@struktur, [])
    end
    
    # strukt: abwechselnd verschachteltes Array: Reihenfolge/Alternativen/Reihenfolge/...
    # pfad_bisher: Array
    def notiere_pfade(strukt, pfad_bisher)      
      key_oder_strukt = strukt.first
      case key_oder_strukt
      when nil then # Ende
        @pfade[pfad_bisher.last] = pfad_bisher
      when Array then
        alternativen = key_oder_strukt
        alternativen.each do |sub_strukt|
          notiere_pfade(sub_strukt, pfad_bisher)
        end
      else
        key = key_oder_strukt 
        notiere_pfade(strukt[1..-1], pfad_bisher + [key])
      end
    end
    
    def pfad_bis(letzter_key)
      @pfade[letzter_key]
    end
  end

  class VObjektBasis < DBF::Record #< ActiveRecord::Base
    def self.vo_sym
      self.name.split("::").last.downcase.to_sym
    end
    def vo_klassen_sym
      self.class.vo_sym
    end
    
    def self.define_attr_method(name, value=nil, &block)
      eigenclass = class << self; self; end
      #eigenclass.send :alias_method, "original_#{name}", name
      eigenclass.send(:undef_method, name) if respond_to? name
      if block_given?
        eigenclass.send :define_method, name, &block
      else
        # use eval instead of a block to work around a memory leak in dev
        # mode in fcgi
        eigenclass.class_eval "def #{name}; #{value.inspect}; end"
      end
    end
    
    def self.belongs_to(*args)
    end
    def self.has_one(*args)
    end
    
    
    def self.has_many(name, options)
      fremd_klassen_name = options[:class_name] || name.capitalize[0..-2]
      
      define_method name do 
        fremd_klasse = eval(fremd_klassen_name)
        fremd_klasse.find(:all, self.primkeyval_hash)
      end
    end
    
    def primkeyval_hash(primkeys = self.class.primary_keys)
      erg = {}
      primkeys.each do |primkey|
        #erg[primkey] = self.send(primkey)
        erg[primkey] = self.attributes[primkey.to_s.downcase]
      end
      erg
    end
    
    def self.set_table_name(neuer_name)
      return if respond_to?(:table_name) and table_name == neuer_name 
      klassenreset      
      define_attr_method :table_name, neuer_name      
    end
    
    def self.klassenreset
      super
      @tabelle = nil
      define_attr_method :table_name, nil
    end
    
    def self.table_name
      nil
    end
    
    def self.connected?
      not self.table_name.nil?
    end

    def self.set_primary_keys(*primkeys)
      primkeys = primkeys.first if primkeys.first.is_a?(Array)
      primkeys = primkeys.map {|pk| pk.to_s.downcase.to_sym}
      sing = class << self; self; end
      #sing.send :alias_method, "original_#{name}", name
      sing.class_eval "def primary_keys; #{primkeys.inspect}; end"
    end
    
    
    def self.tabellen_pfad_kl
      File.join($aktueller_dbf_ordner, table_name + ".dbf")
    end
    def self.tabellen_pfad_gr
      File.join($aktueller_dbf_ordner, table_name + ".DBF")
    end
    
    def self.tabelle
      begin
        @tabelle ||= DBF::Table.new(tabellen_pfad_gr, :record_klasse => self)
      rescue => e
        p e
        if e.class.name =~ /ENOENT/ then
          @tabelle ||= DBF::Table.new(tabellen_pfad_kl, :record_klasse => self)
        else
          raise
        end
      end
    end
    
    def self.find(quanti, options={})
      if not table_name or table_name == "" then
        raise ActiveRecord::ConnectionNotEstablished, "Dbf-Tabelle nicht geöffnet"
      end
      begin
        trc_temp   :find_args, [quanti, options]
        #trc_caller :find_call, 2
        upcase_options = Hash.new
        options.each do |k,v| 
          v = v.strip if v.is_a?(String)
          upcase_options[k.to_s.downcase.to_sym] = v
        end
        tabelle.find quanti, upcase_options
      rescue => e 
        if e.class.name =~ /ENOENT/ then
          raise RuntimeError, "Tabelle nicht gefunden: " + e.message
        else
          raise
        end
      end
    end
    
    def self.name_klein
      self.name.split("::").last.downcase
    end
    
  end

  class VObjektKaskadiert < VObjektBasis
  end


  class Vd < VObjektKaskadiert
    set_primary_keys :vsnr
    #has_one :st, :foreign_key => :vsnr
    #has_one :vp, :foreign_key => :vsnr
    has_many :unsortierte_vks, :class_name => "Vk" ,:foreign_key => :vsnr

    def vk_haupt
      vks.first
    end

private
#    def unsortierte_vks
 #     (@unsortierte_vks = get_vks).each {|vk| vk.vd = self} unless @unsortierte_vks
  #    @unsortierte_vks
   # end

  end


  class St < VObjektKaskadiert  # ZenTest SKIP
    set_primary_keys :vsnr
    belongs_to :vd#, :foreign_key => :vsnr
  end

  class Vp < VObjektKaskadiert  # ZenTest SKIP
    set_primary_keys "vsnr"
    belongs_to :vd#, :foreign_key => :vsnr
  end

  class Vk < VObjektKaskadiert  # ZenTest SKIP
    belongs_to :get_vd, :class_name => "Vd", :foreign_key => :vsnr
    has_many :vvs, :class_name => "Vv" , :foreign_key => [:vsnr, :komp]
    has_many :vts, :class_name => "Vt" , :foreign_key => [:vsnr, :komp] #, :conditions => "vtBAT7V4.komp = vkBAT7V4.komp"
    set_primary_keys :vsnr, :komp

    
    def vv_max
      @max_vj ||= begin
        my_vvs = vvs.to_a # falls kein EIGENES_AR_CACHING
        max_vj = my_vvs.map{|vv| vv.vj}.max
        my_vvs.find {|vv| max_vj == vv.vj }
      end
    end

    def vv_fuer(vj_oder_zgper)
      case vj_oder_zgper
      when Zeit then
        vvs.to_a.find {|vv| vj_oder_zgper == vv.beg_zeit }
      when Integer
        vvs.to_a.find {|vv| vj_oder_zgper == vv.vj }
      else
        trc_aktuellen_error "keine Zeitangabe: #{vj_oder_zgper.inspect}"
      end
    end


    def vt1
      vts.first
    end
  end

  class Vt < VObjektKaskadiert  
    belongs_to :get_vk, :class_name => "Vk", :foreign_key => [:vsnr, :komp]
    set_primary_keys [:vsnr, :komp, :vtnr]

    has_many :get_rks, :class_name => "Rk", :foreign_key => [:vsnr, :komp, :vtnr]
    def rks
      if EIGENES_AR_CACHING then
        (@rks = get_rks).each {|rk| rk.vt = self} unless @rks
        @rks
      else
        trc_temp :vor_getrks
        a = get_rks
        trc_temp :nach_getrks
        a
      end
    end

    has_many :get_vbs, :class_name => "Vb", :foreign_key => [:vsnr, :komp, :vtnr]
    def vbs
      if EIGENES_AR_CACHING then
        (@vbs = get_vbs).each {|vb_1| vb_1.vt = self} unless @vbs
        @vbs
      else
        trc_temp :vor_getvbs
        a = get_vbs
        trc_temp :nach_getvbs
        a
      end
    end

    has_many :get_vfs, :class_name => "Vf", :foreign_key => [:vsnr, :komp, :vtnr]
    def vfs
      if EIGENES_AR_CACHING then
        (@vfs = get_vfs).each {|vf_1| vf_1.vt = self} unless @vfs
        @vfs
      else
        trc_temp :vor_getvfs
        a = get_vfs
        trc_temp :nach_getvfs
        a
      end
    end

    def va
      return @va if @va
      @va = begin
        Va.find( :first, 
                   :vsnr  => vd.vsnr, 
                   :gv    => vk.gv, 
                   :tarif => tarif, 
                   :lfdkz => lfdkz
        )
      rescue ActiveRecord::RecordNotFound
        nil
      end
    end

    def lfdkz
      if tarif =~ /^FB([EL]).$/ then
        $1
      else
        vk.vd.beiart > 0 ? "L" : "E"
      end
    end

    def beg_zeit
      Zeit.jm(begj, begm)
    end
    
    def rel_fkabw
      (vd.fkabw - vt.begm) % 12
    end
    
    def voriger_rpkt(t)
      if t.m > self.rel_fkabw then
        Zeit.jm(t.j, rel_fkabw)
      else
        if t.j > 0 then
          Zeit.jm(t.j - 1, rel_fkabw)
        else
          Zeit.jm(0,0)
        end
      end
    end
        
      

    def rk_nach_zeit(zeit_relativ)
      if not defined? @rk_nach_zeit
        init_verbindung_rk
      end
      erg = @rk_nach_zeit[zeit_relativ]
      if !erg
        trc_hinweis "rk_nach_zeit ist nil bei zeit=", zeit_relativ
        #trc_info "@rk_nach_zeit", @rk_nach_zeit
      end
      erg
    end

    def init_verbindung_rk
      @rk_nach_zeit = {}
      trc_info "vt-rksnachzeit rks-length:",rks.length
      rks.each_with_index {|rk, idx|
        rk.vt_index = idx
        @rk_nach_zeit[z=Zeit.jm(rk.vj, rk.vm)] = rk;
        trc_temp "z#{z}"
      }
      #init_verbindung_vb
    end

    def init_verbindung_vb
      vbs.each do |vb|
        begin
          rk = @rk_nach_zeit[Zeit.jm(vb.vj,vb.vm)]
          rk.vb = vb
          vb.rk = rk
        rescue
          trc_aktuellen_error :verb_vb, 6
        end
      end
    end

  end

  class Rk < VObjektKaskadiert  
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]


    attr_writer :vt_index
    def vt_index
      vt.init_verbindung_rk if not @vt_index
      @vt_index
    end

  end

  class Vb < VObjektKaskadiert  
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
#    attr_accessor :vt_index

  end


  class Vv < VObjektKaskadiert  
    belongs_to :vk, :foreign_key => [ :vsnr, :komp]
  #  has_many :rks
    set_primary_keys [ :vsnr, :komp, :vj]

    def beg_zeit
      Zeit.jm(begj, begm)
    end
    
  end

  class Va < VObjektBasis  
    has_many :get_vts, :class_name => "Vt", :foreign_key => [:vsnr, :tarif ]
    set_primary_keys [:vsnr, :gv, :tarif, :lfdkz]

  end

  class Vf < VObjektKaskadiert 
    belongs_to :get_vt, :class_name => "Vt", :foreign_key => [:vsnr, :komp, :vtnr]
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
#    attr_accessor :vt_index

    
    def beg_zeit
      Zeit.jm(begj, begm)
    end

  end

  class Gz < VObjektBasis  
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
 end

  class Zs < VObjektBasis  
    set_primary_keys [:vsnr, :komp, :vtnr, "vj", "vm"]
  end

  
  VO_KLASSEN = [Vd, St, Vp, Vk, Vt, Rk, Vv, Va, Vb, Vf, Gz, Zs]
  
  VO_KLASSEN.each do |vklasse1|
    trc_temp "vk1##", vklasse1.primary_keys
    VO_KLASSEN.each do |vklasse2|
      if (vklasse2.primary_keys - vklasse1.primary_keys).empty? then
        #trc_temp "vk2->", [vklasse2, vklasse2.name_klein, vklasse2.primary_keys]
        vklasse1.send(:define_method, vklasse2.name_klein) do
          vklasse2.find(:first, self.primkeyval_hash(vklasse2.primary_keys))
        end
      else
        #trc_temp "vk2xx", vklasse2.primary_keys
      end
    end
  end
  #KEY_SYSTEM = KeySystem.new [:vsnr, :komp, :vtnr]
  
end

end # if not defined? DbfDat::VObjektBasis 

if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
