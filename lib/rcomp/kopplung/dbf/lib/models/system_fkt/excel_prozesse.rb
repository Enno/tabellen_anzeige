
module ExcelProzesse
   module_function 

  def alle_laufenden
    zeilen = %x(TASKLIST /FI "IMAGENAME eq EXCEL.EXE" /v 2>NUL)
    zeilen.split("\n").map do |zeile|
      if zeile =~ /^ *EXCEL.EXE\s+(\d+)/ then
        $1.to_i
      end
    end.compact
  end
  #module_function :alle_laufenden
  
  def beende_alle
    alle_laufenden.each do |pid|
      Process.kill("KILL", pid)
    end    
  end
  #module_function :beende_alle
  
end

