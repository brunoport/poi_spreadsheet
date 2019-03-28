require 'rjb'

class PoiSpreadsheet

  def self.init(max_heap)
    Rjb.load(jar_paths, ["-Xmx#{max_heap}"])
  end

  def self.jar_paths
    file_path = File.dirname(__FILE__)

    Dir.glob("#{file_path+'/../apache/'}*.jar").join(':')
  end

  def self.cell_class; load_class('org.apache.poi.ss.usermodel.CellType') end

  def self.load_class(name)
    Rjb.classes[name] || Rjb.import(name)
  end

  def self.load(file, sheet_name=nil, max_heap='1024M')
    init(max_heap)

    Workbook.load(file, sheet_name)
  end

  class Workbook

    attr_accessor :j_book

    def self.load(file, sheet_name=nil)
      book = new

      workbook = ::PoiSpreadsheet.load_class('org.apache.poi.xssf.usermodel.XSSFWorkbook').new(file)

      book.j_book = ::PoiSpreadsheet.load_class('org.apache.poi.xssf.streaming.SXSSFWorkbook').new(workbook, 1, true, false)
      book.sheets(sheet_name)

      book
    end

    def initialize
      @sheets = nil
    end

    # Get sheets
    def sheets(sheet_name=nil)
      @sheets ||= begin
        sheets = {}
        j_book.getNumberOfSheets.times { |i|
          name = j_book.getSheetName(i)

          next if sheet_name && name != sheet_name

          j_sheet = j_book.getSheetAt(i)
          sheet = Worksheet.from_sheet(j_sheet)
          sheet.book = self
          sheets[name] = sheet
        }
        sheets
      end
    end

    def create_sheet(name)
      sheets << Worksheet.from_sheet(j_book.createSheet(name))
    end

    def clone_sheet(index)
      sheets << Worksheet.from_sheet(j_book.cloneSheet(index))
    end

    def remove_sheet_at index
      j_book.removeSheetAt(index)
      @sheets.delete_at(index)
    end

    # Get sheet by name
    def [](k)
      sheets[k]
    end

    def save(file_name)
      @file_output_class ||= (Rjb.classes['java.io.FileOutputStream'] || Rjb.import('java.io.FileOutputStream'))
      out = @file_output_class.new(file_name)

      begin
        j_book.write(out)
      ensure
        out.close
        j_book.close

        self.j_book = nil
      end
    end
  end

  class Worksheet

    attr_accessor :j_sheet
    attr_accessor :book

    def initialize
      @rows = {}
    end

    def set_values(row, start_col, values)
      j_row = j_sheet.createRow(row)
      values.map { |v| fix_value_to_java(v) }
            .each_with_index { |v, col| j_row.createCell(start_col + col).setCellValue(v) }
    end

    def fix_value_to_java(v)
      v.is_a?(Integer) ? fix_integer_value(v) : v
    end

    def fix_integer_value(v)
      num_bytes = [42].pack('i').size
      num_bits = num_bytes * 8
      max = 2**(num_bits - 2) - 1

      (v > max) ? v.to_s : v
    end

    # get cell
    def [](row)
      j_row = j_sheet.getRow(row) || j_sheet.createRow(row)
      row = Row.from_row(j_row)
      row
    end

    def self.from_sheet j_sheet
      sheet = new
      sheet.j_sheet = j_sheet
      sheet
    end

    def name
      j_sheet.getSheetName
    end

    def name= name
      j_book = j_sheet.getWorkbook
      j_book.setSheetName(j_book.getSheetIndex(j_sheet), name)
    end

    class Row

      attr_accessor :j_row
      attr_accessor :sheet

      def self.symbol_type(constant)
        @types ||= begin
          cell = ::PoiSpreadsheet.cell_class
          {
            cell.BOOLEAN.toString => :boolean,
            cell.NUMERIC.toString => :numeric,
            cell.STRING.toString => :string,
            cell.BLANK.toString => :blank,
            cell.ERROR.toString => :error,
            cell.FORMULA.toString => :formula
          }
        end
        @types[constant.toString]
      end

      def []= col, value
        cell = j_row.getCell(col) || j_row.createCell(col)
        cell.setCellValue(value)
      end

      def [] col
        unless cell = j_row.getCell(col)
          return nil
        end

        type = self.class.symbol_type(cell.getCellType())

        case type
        when :boolean
          cell.getBooleanCellValue()
        when :numeric
          cell.getNumericCellValue()
        when :string
          cell.getStringCellValue()
        when :blank
          nil
        when :error
          cell.getErrorCellValue()
        when :formula
          cell.getNumericCellValue()
        end
      end

      def self.from_row j_row
        row = new
        row.j_row = j_row
        row
      end

    end

  end

end
