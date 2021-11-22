require 'roo'
# require 'byebug'


# $databasexlfilename = "Kooperacija1.xlsx" #give your database file name
$databasexlfilename = "Input_output1.xlsx" #give your database file name

xls_file = Roo::Excelx.new($databasexlfilename)

# sheet = xls_fil e.sheets.first

# csv = Roo::CSV.new(xls_file.sheets.to_csv)
xls_file.sheets.each { |s| 
  xls_file.default_sheet = s
    xls_file.to_csv("#{s}.csv")
  }