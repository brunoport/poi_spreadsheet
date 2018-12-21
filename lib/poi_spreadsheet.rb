require 'rjb'

class PoiSpreadsheet

  def self.init
    apache_poi_path = File.dirname(__FILE__)+'/../apache/poi-4.0.1.jar'
    Rjb::load(apache_poi_path, ['-Xmx512M'])

    Rjb::add_jar(File.dirname(__FILE__)+'/../apache/commons-collections4-4.2.jar')
    Rjb::add_jar(File.dirname(__FILE__)+'/../apache/xmlbeans-3.0.2.jar')
    Rjb::add_jar(File.dirname(__FILE__)+'/../apache/commons-compress-1.18.jar')
    Rjb::add_jar(File.dirname(__FILE__)+'/../apache/poi-ooxml-schemas-4.0.1.jar')
    Rjb::add_jar(File.dirname(__FILE__)+'/../apache/poi-ooxml-4.0.1.jar')

    @cell_class = Rjb::import('org.apache.poi.ss.usermodel.CellType')

    # You can import all java classes that you need
    @loaded = true
  end

  def self.cell_class; @cell_class; end

  def self.load(file, sheet_name=nil)
    unless @loaded
      init
    end
    Workbook.load file, sheet_name
  end


  class Workbook

    attr_accessor :j_book

    def self.load(file, sheet_name=nil)
      @file_name = file

      @workbook_class = Rjb::import('org.apache.poi.xssf.usermodel.XSSFWorkbook')
      @file_input_class = Rjb::import('java.io.File')
      @zip_secure_file_class = Rjb::import('org.apache.poi.openxml4j.util.ZipSecureFile')

      @file_input = @file_input_class.new(file)

      @zip_secure_file_class.setMinInflateRatio(0);

      book = new
      @sworkbook_class = Rjb::import('org.apache.poi.xssf.streaming.SXSSFWorkbook')

      book.j_book = @sworkbook_class.new(@workbook_class.new(@file_input), 1, true, true)
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
        self.j_book.getNumberOfSheets.times { |i|
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

    def create_sheet name
      self.sheets << Worksheet.from_sheet(j_book.createSheet(name))
    end

    def clone_sheet index
      self.sheets << Worksheet.from_sheet(j_book.cloneSheet(index))
    end

    def remove_sheet_at index
      j_book.removeSheetAt(index)
      @sheets.delete_at(index)
    end

    # Get sheet by name
    def [](k)
      sheets[k]
    end

    def save file_name = @file_name
      @file_output_class ||= Rjb::import('java.io.FileOutputStream')
      out = @file_output_class.new(file_name);

      begin
        j_book.write(out)
      ensure
        out.close
        j_book.dispose
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
      values.each_with_index { |v, col| j_row.createCell(start_col + col).setCellValue(v) }
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
