# encoxxxxxxxding: us-ascii
# encodffffing: utf-16le

TAFELN_DIR = File.dirname(File.expand_path(__FILE__).sub("/lib",""))

require 'spreadsheet'


#f = File.open(TAFELN_DIR+'/daten/ExstarPD.xls')
#erg = f.readpartial(90)
#f.close
#puts erg
#$stdout.flush

#Spreadsheet.client_encoding =  'UTF-16LE' #  'ASCII-8BIT' #'UTF-8' 'US-ASCII' #
#Spreadsheet.client_encoding = 'US-ASCII' #  'ASCII-8BIT' #'UTF-8' 'US-ASCII' #
#Spreadsheet.client_encoding = 'ASCII-8BIT' #'UTF-8' 'US-ASCII' #
Spreadsheet.client_encoding = 'UTF-8'

book = Spreadsheet.open TAFELN_DIR+'/daten/ExstarPD.xls'
# We can either access all the Worksheets in a Workbook?

book.worksheets

#?or access them by index or name (encoded in your client_encoding)

  sheet1 = book.worksheet 0
  #sheet2 = Book.worksheet 'Sheet1'

# Now you can either iterate over all rows that contain some data. A call to Worksheet.each without argument will omit empty rows at the beginning of the Worksheet:

i = 0
sheet1.each do |row|
    break if (i+=1) > 22
    str = row.to_a #.map{|v|String===v ? v.encode("ASCII-8BIT") : v}
    #p str
    p case str
      when String
        str.encode "ASCII-8BIT"
    else
      str
    end
    # do something interesting with a row
  end

# Or you can tell Worksheet how many rows should be omitted at the beginning. The following starts at the 3rd row, regardless of whether or not it or the preceding rows contain any data:

  #sheet2.each 2 do |row|
    # do something interesting with a row
  #end
