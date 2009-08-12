# To change this template, choose Tools | Templates
# and open the template in the editor.

class DatenModell <  javax.swing.table.AbstractTableModel
  def initialize
    super
    @daten = Array.new(7) {|i| Array.new(4,nil)}
  end

  def getRowCount
    7
  end

  def getColumnCount
    4
  end

  def getColumnName(j)
    (j*j).to_s
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
