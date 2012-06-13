# encoding: utf-8
#require "torg12_xls_to_xml/version"
require 'spreadsheet'
require "nokogiri"

module Torg12XlsToXml

  def self.perform
    files_path = ENV['HOME']+"/torg12/"
    Dir.glob(File.join(files_path , '*.xls')).each do |file|
       convert(file)
       write_xml
       file_xml = file.chomp(File.extname(file)) + ".xml"
       write_file(file_xml)
    end
  end

  def self.write_file(file)
    #$nad = "we"
    #puts builder.to_xml
    f = File.new(file, 'w')
    #File.open('file.xml', 'w'){ |file| file.write string }
    f.write(@builder.to_xml)
    f.close
    #f.save
    #we = "фв"
  end
  def self.convert(file_to_convert)
    #read file
    #@load = Xl.find(params[:id])
    #@list= 'public'+@load.xml.url
    @spreed = Spreadsheet.open file_to_convert
    $sheet = @spreed.worksheet 0
    #search the file
    $test = 12
    $product_tax = []
    $product_sum = []
    $product_amount = []
    $product_name = []
    $product_cod = []
    $product_named = []
    #$product_wiegth = []
    $product_rate = []

    @start_product = []
    step = 0
    if !$sheet.nil?
      $sheet.each_with_index do |rows, index_row|
        rows.each_with_index do |cells, index_cell|
          if cells.to_s == "Поставщик"
            @contractor_1_row = index_row
            @contractor_1_coll = index_cell + 1
          end
          if cells.to_s == "Плательщик"
            @contractor_2_row = index_row
            @contractor_2_coll = index_cell + 1
          end
          if cells.to_s == "наименование, характеристика, сорт, артикул товара"
            @name_coll = index_cell
          end
          if cells.to_s == "код по ОКЕИ"
            @cod_coll = index_cell
          end
          if cells.to_s == "наиме- нование"
            @named_coll = index_cell
          end
          if !cells.to_s["нетто"].nil? and !cells.to_s["Коли"].nil?
            @amount_coll = index_cell
          end
          #if cells.to_s == "Масса брутто"
          #  @wieght_coll = index_cell + 1
          #end
          if !cells.to_s["Сумма с"].nil?
            @sum_coll = index_cell
          end
          if cells.to_s == "ставка, %"
            @rate_coll = index_cell
          end
          if !cells.to_s["сумма, "].nil?
            @tax_coll = index_cell
          end
          @margin_down = 3
          if cells.to_s == "Товар"
            @start_product[step] = index_row.to_i
            @start_product[step] += @margin_down
            step += 1
          end
          if !cells.to_s["Всего по накладной"].nil?
            @sum_all_row = index_row
          end

        end
      end
    end
    j = 0
    @start_product.each do |s|
      i = 0
      begin
        $product_name[j] = findcell(s + i, @name_coll.to_i)
        $product_cod[j] = findcell(s + i, @cod_coll.to_i)
        $product_named[j] = findcell(s + i, @named_coll.to_i)
        #$product_wiegth[i] = findcell(@start_product+i, @wieght_coll)
        $product_amount[j] = findcell(s + i, @amount_coll.to_i)
        $product_sum[j] = findcell(s + i, @sum_coll.to_i)
        $product_tax[j] = findcell(s + i, @tax_coll.to_i)
        $product_rate[j] = findcell(s + i, @rate_coll.to_i).to_s + "%"
        i += 1
        j += 1
      end  while !(findcell(s + i,@name_coll.to_i).nil?)
    end

    $contractor_1 = findcell(@contractor_1_row,@contractor_1_coll).split(',')
    $contractor_2 = findcell(@contractor_2_row,@contractor_2_coll).split(',')
    if !findcell(@contractor_2_row,@contractor_2_coll).to_s["ИНН"].nil?
      $inn_2 = $contractor_2[1].split(" ")[1]
    else
      $inn_2 = "0000000000"
    end

    $product_all_sum = findcell(@sum_all_row, @sum_coll)


    #$test = $product_name

  end

 def self.write_xml
    @builder = Nokogiri::XML::Builder.new(:encoding => 'UTF-16') do
    КоммерческаяИнформация('ВерсияСхемы' => '2.03', 'ДатаФормирования' => DateTime.now.to_s) {
      Документ {
        Ид
        Номер
        Дата
        ХозОперация "Отпуск товара"
        Роль
        Валюта "руб"
        Курс "1"
        Сумма $product_all_sum
        Контрагенты {
          Контрагент {
            Ид
            Наименование $contractor_1[0]
            ОфициальноеНаименование $contractor_1[0]
            ЮридическийАдрес {
              Представление $contractor_1[2] + ", " + $contractor_1[3] + ", " + $contractor_1[4] + ", " + $contractor_1[5] + ", " + $contractor_1[6] + ", " + $contractor_1[7]
              АдресноеПоле {
                Тип "Почтовый индекс"
                Значение $contractor_1[2]
              }
              АдресноеПоле {
                Тип "Регион"
                Значение $contractor_1[3]
              }
              АдресноеПоле {
                Тип "Населенный пункт"
                Значение $contractor_1[4]
              }
              АдресноеПоле {
                Тип "Улица"
                Значение $contractor_1[5]
              }
              АдресноеПоле {
                Тип "Дом"
                Значение $contractor_1[6].split(" ")[2]
              }
              АдресноеПоле {
                Тип "Квартира"
                Значение $contractor_1[7].split(/[. ]+/)[2]
              }
            }
            ИНН $contractor_1[1].split(" ")[1]
            КПП
            Роль "Продавец"
          }
          Контрагент {
            ИД
            Наименование $contractor_2[0]
            ПолноеНаименование $contractor_2[0]
            ИНН $inn_2
            Роль "Покупатель"
          }
        }
        Время
        Налоги {
          Налог {
            Наименовение "НДС"
            УчтеноВСумме
            Сумма "----после всего остального----"
          }
        }
        Товары {
          j = 0
          while !$product_name[j].nil?
          Товар {
            Ид
            Артикул
            Наименование $product_name[j]
            БазоваяЕденица("Код" => $product_cod[j],  "НаименованиеПолное" => $product_named[j]) {
              Пересчет {
                Еденица $product_named[j]
                Коэффициент
                ДополнительныеДанные {
                  ЗначениеРеквизита {
                    Наименование "Вес"
                    Значение 0#$product_wiegth[j]
                  }
                  ЗначениеРеквизита {
                    Наименование "Объем"
                    Значение 0
                  }
                }
              }
            }
            Группы {
              Ид
            }
            СтавкиНалогов {
              СтавкиНалога {
                Наименование "НДС"
                Ставка $product_rate[j]
              }
            }
            ЗначениеРеквизитов {
              ЗначениеРеквизита {
                Наименование "ТипНоменклатуры"
                Значение "Товар"
              }
              ЗначениеРеквизита {
                Наименование "ТипНоменклатуры"
                Значение "Товар"
              }
              ЗначениеРеквизита {
                Наименование "НаименованиеКраткое"
                Значение $product_name[j]
              }
              ЗначениеРеквизита {
                Наименование "НаименованиеПолное"
                Значение $product_name[j]
              }
            }
            ИдКаталога
            ЦенаЗаЕдиницу ($product_sum[j].to_i/$product_amount[j].to_i).to_i
            Количество $product_amount[j]
            Сумма $product_sum[j]
            Единица $product_named[j]
            Коэффициент
            Наголи {
              Налог {
                Наименование
                УчтеноВСумме "true"
                Сумма $product_tax[j]
              }

            }

          }
            j+=1
          end
        }
      }
    }
    end



  end
 def self.findcell(row,coll)
    @findrow = $sheet.row(row)
    @findcell = @findrow[coll]
  end

end


Torg12XlsToXml.perform