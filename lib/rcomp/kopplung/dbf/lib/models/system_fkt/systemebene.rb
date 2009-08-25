
if not defined?(InifileHash) then

require 'schmiedebasis'

# Inspired by Yukimi_Sake-san's inifile.rb.
class InifileHash < Hash
=begin
===Methods
---deletekey(section,key)
  Delete ((|key|)) and associated value in ((|section|)).
---erasesection(section)
  Erase ((|section|)) and associated keys and values.
---read( section, key, [default=""])
  Read value associated ((|key|)) in ((|section|)).
  If ((|section|)) or ((|key|)) not exists ,value uses ((|default|)).
---write( section, key, value)
  Write ((|value|)) in assiciated ((|key|)) and ((|section|)).
  If ((|section|)) or ((|key|)) not exists , make them.
---frash
  Write into ((|inifilename|)) of instance.
=end

  # inkonsequent: Abschnittsname als String, Schlüssel als Symbol
  def initialize(ini_datei_name, &blk)
    @proc_neuer_abschnitt = blk || proc{Hash.new}
    abschnittsname=""
    @fn = ini_datei_name
    if File.exist?(@fn) then
      f=open @fn
      f.each do |t|
        if t =~ /\[(.+)\]/ then
          abschnittsname = $1.strip
          self[abschnittsname] = @proc_neuer_abschnitt.call(abschnittsname)
        elsif t =~/.+=/ then
          key, value =t.split(/=/)
          value.strip!
          case value
            when "true"   then value = true
            when "false"  then value = false
            when /^[-+\d]+$/ then value = value.to_i
          end

          self[abschnittsname][key.strip.to_sym] = value
        end
      end
      f.close
    end
  end

  def [](aKey)
    fetch(aKey, &@proc_neuer_abschnitt)
  end

  def deletekey(abschnittsname,key)
    self[abschnittsname.strip].delete(key)
  end

  def erasesection(abschnittsname)
    self.delete(abschnittsname)
  end

  def read( abschnittsname, key, default="")
    if self[abschnittsname] && r=self[abschnittsname][key] then r else default end
  end

  def write( abschnittsname, key, value)
    self.update({abschnittsname.strip=>{}}) if self[abschnittsname.strip] == nil
    self[abschnittsname.strip].update({key.strip => (value.to_s).strip})
  end

  def flush
    open(@fn,"w") { |f|
      self.sort.each do |(abschnittsname, inhalt)|
        f.write "[#{abschnittsname}]\n"
        inhalt.sort.each do |(k1,v1)|
           f.write "#{k1}=#{v1}\n"
        end
        f.write "\n"
      end
    }
  end

end

end # unless defined?(InifileHash)

if __FILE__ == $0
  durchlaufe_unittests($0)
end

