# -*- coding: utf-8 -*-"

# encoding: utf-8

class DatenModellDummy <  javax.swing.table.AbstractTableModel

  def initialize()
    super()
    #@daten = Array.new(7) {|i| Array.new(4, nil)}
    @daten = [
      [1,  2, 3,  4.5, nil],
      [21,22,23, 24.2, nil],
      [31,32,33,  "a", nil],
      [41,42,43, 44.4, nil],
      [51,52,53, 54.5, nil]
      ]
      p @daten
  end

  def getRowCount
    @daten.size
  end

  def getColumnCount
    5
  end

  def getColumnName(j)
    j*j
  end

  def getValueAt(i, j)
    @daten[i][j].to_s if @daten[i][j]
  end

  def setValueAt(obj, i, j)
    @daten[i][j] = obj
  end

  def isCellEditable(i, j)
    true
  end

end
