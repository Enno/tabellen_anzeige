# encoding: utf-8
TAFELN_DIR = File.dirname(File.expand_path(__FILE__))

require 'spreadsheet'
require TAFELN_DIR + '/nx_excel'
require TAFELN_DIR + '/biometrie_db'

Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open File.dirname(TAFELN_DIR)+'/daten/ExstarPD.xls'
blatt1 = book.worksheet "BiomVekt"

p :blatt_geoeffnet
a = Time.now
i = 0
arr = Array.new(250,nil)
zeilen = blatt1.map do |zeile|
  (zeile.to_a + arr).first(250)
end
b = Time.now
spalten = zeilen.transpose
c = Time.now
biom_vekt = {}

spalten.each do |id, *rest|
  BIOMETRIE_DB.nx_quellen[id] = NxExcel.new(rest)
end
d = Time.now
p [b-a, c-b, d-c]
p spalten[4]