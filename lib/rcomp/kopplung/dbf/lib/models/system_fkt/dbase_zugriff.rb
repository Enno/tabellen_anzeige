# encoding: iso-8859-15


PROGRAMM_MIT_DBF_ZUGRIFF = true unless defined? PROGRAMM_MIT_DBF_ZUGRIFF

if PROGRAMM_MIT_DBF_ZUGRIFF then

require 'schmiedebasis'

trc_info "dbszugr_aufrufer", caller

case :d
when :d then require 'models/system_fkt/dbase_zugriff_direkt'
when :a then require 'models/system_fkt/dbase_zugriff_apollo'
end

end


if __FILE__ == $0 then
  durchlaufe_unittests($0)
end
