#!ruby
# encoding: cp932

# @author yuzukinhansyoku
# @date 2012/7/8
# 
# 
# �w�肳�ꂽ�f�B���N�g�����ɂ��� Excel �t�@�C�����ׂĂ̑S�V�[�g��
# csv �t�@�C���Ƃ��ĕۑ�����B
# 
# Excel �t�@�C������������f�B���N�g���ւ̃p�X
# csv �t�@�C���������o���f�B���N�g���ւ̃p�X
# 
# csv �t�@�C������
# 	{Excel �t�@�C����}_{�V�[�g��}.csv �ƂȂ�܂��B
# 		sample1.xlsx �� Sheet1 �����Ƃɂ����ꍇ�Asample1_Sheet1.csv �ƂȂ�܂��B
# 
# How to use.
# 	ruby excel_sheets_to_csv.rb .\sources .\out


require 'win32OLE'
require 'find'
require 'fileutils'
require 'pp'
require 'csv'

module Excel
end

class OneSheet
	def setup(name, datas)
		@name = name
		if datas.instance_of?(Array)
			@datas = datas
		else
			@datas = [[datas]]
		end
	end
	attr_reader :name
	attr_reader :datas
end

def output(src_dir, out_dir)
	p 'output'
	p Dir::pwd
	FileUtils.mkdir_p File::absolute_path(out_dir)
	
	excel = WIN32OLE.new("Excel.Application")
	WIN32OLE.const_load(excel, Excel)
	
	Find::find(File::absolute_path(src_dir)) {|f|
		if f.to_s =~ /.+\.xlsx|.+\.xlsm|.+\.xlsb|.+\.xls/
			workbook = excel.workbooks.open(File::absolute_path(f.to_s.gsub('/', '\\')), 'ReadOnly'=>true)
			begin
				bookname = workbook.name.gsub(/\..+/, '')
				p bookname
				sheets = []
				workbook.worksheets.each {|sheet|
					p "  #{sheet.name}"
					os = OneSheet.new
					os.setup(sheet.name, sheet.usedrange.value)
					sheets << os
				}
				od = File::absolute_path("#{File::absolute_path(out_dir)}\\")
				sheets.collect {|s|
					CSV.open("#{od}\\#{bookname}_#{s.name}.csv", "w") {|f|
						s.datas.each {|row|
							f << row
						}
					}
				}
			ensure
				workbook.close('SaveChanges'=>false)
			end
			
		end
	}
	excel.quit
end

def main
	p 'main'
	if ARGV.size == 2
		source_directory = ARGV[0]
		out_directory = ARGV[1]
		output(source_directory, out_directory)
	end
end



if $0 == __FILE__
	main
end





























