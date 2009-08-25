# -*- coding: utf-8 -*-"

# encoding: utf-8

#require 'ffmath'


class DatenModell <  javax.swing.table.AbstractTableModel

  def initialize(rechen_zeilen)
    super()
    #super
    @rechen_zeilen = rechen_zeilen
    @tabellen_klasse = @rechen_zeilen.first.class
    @spalten_keys = @tabellen_klasse.send("_ausgabe_groeszen")
    p [:@spalten_keys, @spalten_keys]
  end

  def getRowCount
    @rechen_zeilen.size
  end

  def getColumnCount
    @spalten_keys.size
  end

  def getColumnName(j)
    @spalten_keys[j].to_s
  end

  def getValueAt(i, j)
    zeile = @rechen_zeilen[i]
    zeile.send(@spalten_keys[j]).to_s
  end

  def setValueAt(obj, i, j)
  end

  def isCellEditable(i, j)
    #true
    false
  end

end

class DatenModellDummy <  javax.swing.table.AbstractTableModel

  def initialize()
    super()
    @daten = Array.new(7) {|i| Array.new(4, nil)}
  end

  def getRowCount
    7
  end

  def getColumnCount
    4
  end

  def getColumnName(j)
    j*j
  end

  def getValueAt(i, j)
    @daten[i][j]
  end

  def setValueAt(obj, i, j)
    @daten[i][j] = obj
  end

  def isCellEditable(i, j)
    true
  end
end
