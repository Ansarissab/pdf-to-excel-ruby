require "pdf-reader"
require 'write_xlsx'
require 'spreadsheet'
require 'roo'
require 'byebug'

$databasexlfilename = "Kooperacija1.xlsx" #give your database file name
$inputfilename = "Input.pdf" #give your input file name
$outputfilename = "#{$inputfilename.gsub(".pdf","")}_output1.xlsx"  #give expected output file name

def getdatabasexl
  xlsx = Roo::Excelx.new($databasexlfilename)
  forcounter = 1
  xlsx.each_row_streaming do |row|
    if forcounter>1
      itemno = row[1].value rescue ""
      mat = row[4].value.to_s.split('.')[0] rescue ""
      mat = Nokogiri::HTML(mat).text rescue ""
      title = row[5].value.strip rescue ""
      title = Nokogiri::HTML(title).text rescue ""
      a = row[6].value.round(2) rescue ""
      gfor = row[6].formula rescue ""
      b = row[7].value.round(2) rescue ""
      hfor = row[7].formula rescue ""
      c = row[8].value rescue ""
      d = row[9].value rescue ""
      e = row[10].value.round(2) rescue ""
      kfor = row[10].formula rescue ""
      f = row[11].value.round(2) rescue ""
      lfor = row[11].formula rescue ""
      g = row[12].value.round(2) rescue ""
      mfor = row[12].formula rescue ""
      pov = row[13].value.strip rescue ""
      tez = row[14].value rescue ""
      uk = row[15].value.round(3) rescue ""
      pfor = row[15].formula rescue ""
      h = row[16].value rescue ""
      i = row[17].value.round(3) rescue ""
      rfor = row[17].formula rescue ""
      j = row[18].value rescue ""
      k = row[19].value rescue ""
      l = row[20].value rescue ""
      @alldata.push({"itemno":"#{itemno}","title":"#{title}","mat":"#{mat}","pov":"#{pov}","tez":"#{tez}","uk":"#{uk}","a":"#{a}","b":"#{b}","c":"#{c}","d":"#{d}","e":"#{e}","f":"#{f}","g":"#{g}","h":"#{h}","i":"#{i}","j":"#{j}","k":"#{k}","l":"#{l}","gfor":"#{gfor}","hfor":"#{hfor}","kfor":"#{kfor}","lfor":"#{lfor}","mfor":"#{mfor}","pfor":"#{pfor}","rfor":"#{rfor}","rownum":"#{forcounter}"})
    end
  forcounter = forcounter+1
  end
end

def generateoutput
  workbook = WriteXLSX.new($outputfilename)
  worksheet = workbook.add_worksheet
  headerformat = workbook.add_format
  headerformat.set_bold
  headerformat.set_align('center')
  headerformat.set_size(11)
  headerformat.set_font('Arial')
  headerformat.set_border(1)
  globalformat = workbook.add_format
  globalformat.set_size(11)
  globalformat.set_font('Arial')
  globalformat.set_align('center')
  globalformat.set_border(1)
  mainhead = workbook.add_format
  mainhead.set_bold
  mainhead.set_size(14)
  mainhead.set_font('Arial')
  mainhead.set_border(1)
  mainhead.set_left_color('white')
  mainhead.set_right_color('white')
  columnD = workbook.add_format
  columnD.set_bold
  worksheet.set_column('A:A', 10)
  worksheet.set_column('B:B', 20)
  worksheet.set_column('D:D', 20)
  worksheet.set_column('E:E', 30)
  worksheet.set_column('F:F', 40)
  worksheet.set_column('S:S', 40)
  worksheet.set_column('T:T', 40)
  worksheet.set_column('Q:Q', 20)
  worksheet.set_column('U:U', 20)
  currency_format = workbook.add_format({'num_format': '0.0 €'})
  currency_format_2 = workbook.add_format({'num_format': '0.0 "kn"'})
  # maintxt = "Purchase Order No.#{@ponumber}"
  # worksheet.write(2,4, maintxt, mainhead)
  mainhead2 = workbook.add_format
  # mainhead2.set_border(1)
  # mainhead2.set_left_color('white')
  # mainhead2.set_right_color('white')
  worksheet.write(2,5, '', mainhead2)
  worksheet.write(2,6, '', mainhead2)
  worksheet.write(2,7, '', mainhead2)
  format1 = workbook.add_format
  format1.set_size(11)
  format1.set_font('Arial')
  format1.set_align('center')
  format1.set_border(1)
  format1.set_num_format('0.00')
  format2 = workbook.add_format
  format2.set_size(11)
  format2.set_font('Arial')
  format2.set_align('center')
  format2.set_border(1)
  format2.set_num_format('0.000')
  currency_format.set_size(11)
  currency_format.set_font('Arial')
  currency_format.set_align('center')
  currency_format.set_border(1)
  currency_format_2.set_size(11)
  currency_format_2.set_font('Arial')
  currency_format_2.set_align('center')
  currency_format_2.set_border(1)
  header =  [
    ['Po', 'Item no.', 'St.', 'Reference', 'Mat-Nr.', 'Benennung', 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'Pov. zaš.', 'Težina', 'Uk.tež:', 'h', 'i', 'j', 'k', 'l','gfor','hfor','kfor','lfor','mfor','pfor','rfor']
  ]
  worksheet.write_col('A1', header, headerformat)
  @values.each_with_index  do  |v,i|
      worksheet.write("#{v[28]}".to_i-1,0, v[0], globalformat)
      worksheet.write("#{v[28]}".to_i-1,1, v[1], globalformat)
      worksheet.write("#{v[28]}".to_i-1,2, v[2], globalformat)
      worksheet.write("#{v[28]}".to_i-1,3, v[3], globalformat)
      worksheet.write("#{v[28]}".to_i-1,4, v[4], globalformat)
      worksheet.write("#{v[28]}".to_i-1,5, v[5], globalformat)
      worksheet.write("#{v[28]}".to_i-1,6, v[6], globalformat)
      worksheet.write("#{v[28]}".to_i-1,7, v[7], globalformat)
      worksheet.write("#{v[28]}".to_i-1,8, v[8], globalformat)
      worksheet.write("#{v[28]}".to_i-1,9, v[9].to_f,currency_format)
      worksheet.write("#{v[28]}".to_i-1,10, v[10].to_f,currency_format)
      worksheet.write("#{v[28]}".to_i-1,11, v[11].to_f,currency_format_2)
      worksheet.write("#{v[28]}".to_i-1,12, v[12].to_f,currency_format_2)
      worksheet.write("#{v[28]}".to_i-1,13, v[13], globalformat)
      worksheet.write("#{v[28]}".to_i-1,14, v[14], globalformat)
      worksheet.write("#{v[28]}".to_i-1,15, v[15], globalformat)
      worksheet.write("#{v[28]}".to_i-1,16, v[16], globalformat)
      worksheet.write("#{v[28]}".to_i-1,17, v[17], globalformat)
      worksheet.write("#{v[28]}".to_i-1,18, v[18], globalformat)
      worksheet.write("#{v[28]}".to_i-1,19, v[19], globalformat)
      worksheet.write("#{v[28]}".to_i-1,20, v[20], globalformat)
      if v[21]!=""
      worksheet.write_array_formula("G#{v[28].to_i}:G#{v[28].to_i}", "=#{v[21]}",format1)
      worksheet.write("#{v[28]}".to_i-1, 21, v[21], globalformat)
      end
      if v[22]!=""
      worksheet.write_array_formula("H#{v[28].to_i}:H#{v[28].to_i}", "=#{v[22]}",format1)
      worksheet.write("#{v[28]}".to_i-1, 22, v[22], globalformat)
      end
      if v[23]!=""
      worksheet.write_array_formula("K#{v[28].to_i}:K#{v[28].to_i}", "=#{v[23]}",currency_format)
      worksheet.write("#{v[28]}".to_i-1, 23, v[23], globalformat)
      end
      if v[24]!=""
      worksheet.write_array_formula("L#{v[28].to_i}:L#{v[28].to_i}", "=#{v[24]}",currency_format_2)
      worksheet.write("#{v[28]}".to_i-1, 24, v[24], globalformat)
      end
      if v[25]!=""
      worksheet.write_array_formula("M#{v[28].to_i}:M#{v[28].to_i}", "=#{v[25]}",currency_format_2)
      worksheet.write("#{v[28]}".to_i-1, 25, v[25], globalformat)
      end
      if v[26]!=""
      worksheet.write_array_formula("P#{v[28].to_i}:P#{v[28].to_i}", "=#{v[26]}",format2)
      worksheet.write("#{v[28]}".to_i-1, 26, v[26], globalformat)
      end
      if v[27]!=""
      worksheet.write_array_formula("R#{v[28].to_i}:R#{v[28].to_i}", "=#{v[27]}",format2)
      worksheet.write("#{v[28]}".to_i-1, 27, v[27], globalformat)
      end
  end
  workbook.close
end

def getdetails
  @alldata = []
  @values = []
  getdatabasexl #read database file
  reader = PDF::Reader.new($inputfilename)
  reader.pages.each do |page|
    lines = page.text.scan(/^.+/)
    islock1 = 0
    islock2 = 0
    lines.each do |line|
      if (line.include?"PC" and islock1==0)
        @po = line.split(" ")[0] rescue ""
        @itemno = line.split(" ")[1] rescue ""
        @stock = line.split(" ")[2] rescue ""
        islock1 = 1
      end
      if (line.include?"Reference:" and islock2==0)
        @reference = line.split("Reference: ")[1] rescue ""
        islock2 = 1
      end
      if line.include?"Purchase order No.:"
        @ponumber = line.split("Purchase order No.:")[1].split("Purchaser")[0].strip rescue ""
      end
      if line.include?"PO date:"
        @podate = line.split("PO date:")[1].split("Phone")[0].strip rescue ""
      end
      if (islock1==1 and islock2==1)
        @alldata.each do |xldt|
          if xldt[:itemno] == @itemno
            dataeach = ["#{@po}","#{@itemno}","#{@stock}","#{@reference}","#{xldt[:mat]}","#{xldt[:title]}","#{xldt[:a]}","#{xldt[:b]}","#{xldt[:c]}","#{xldt[:d]}","#{xldt[:e]}","#{xldt[:f]}","#{xldt[:g]}","#{xldt[:pov]}","#{xldt[:tez]}","#{xldt[:uk]}","#{xldt[:h]}","#{xldt[:i]}","#{xldt[:j]}","#{xldt[:k]}","#{xldt[:l]}","#{xldt[:gfor]}","#{xldt[:hfor]}","#{xldt[:kfor]}","#{xldt[:lfor]}","#{xldt[:mfor]}","#{xldt[:pfor]}","#{xldt[:rfor]}","#{xldt[:rownum]}"]
            @values.push(dataeach)
          end
        end
        islock2 = 0
        islock1 = 0
        @po = ""
        @itemno = ""
        @stock = ""
        @reference = ""
      end
    end
  end
  generateoutput #write o/p
end

#start
getdetails