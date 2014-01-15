require 'roo'
require 'json'
require 'erb'
require './library/la-finder.rb'

class RdfCreator
  SPREADSHEETS_DIR = "spreadsheets/"
  TEMPLATE_DIR = "templates/"
  CONFIG_DIR = "config/"
  OUTPUT_DIR = "outputs/"

  def initialize(project)
    spreadsheet_name = SPREADSHEETS_DIR + project + ".xlsx"
    config_name = CONFIG_DIR + project + ".conf"


    # open files
    @spreadsheet = Excelx.new(spreadsheet_name)
    config = File.new(config_name,'r')

    
    # read config
    @conf_array = JSON.parse(config.read)
    

    
  end

  # takes a spreadsheet cell and returns a suitable literal - string, int, decimal (in future could add date)
  # what about the case when something should be a float but happens to be 123.0 (ie 0 after decimal point)?
  def literal(row,col)
    val = @spreadsheet.cell(row,col)
    celltype = @spreadsheet.celltype(row,col)
    if celltype == :string
      lit = "\"#{val}\""
    elsif celltype == :float 
      if val.to_i.to_f == val
        lit = "\"#{val.to_i}\"^^<http://www.w3.org/2001/XMLSchema#integer>"
      else
        lit = "\"#{val}\"^^<http://www.w3.org/2001/XMLSchema#decimal>"
      end
    else
      raise "Unexpected cell type #{celltype} #{row} #{col} #{celltype.class}"
    end
    return lit
    
  end

# process config file
  def setup(conf)
    ranges = []

    conf.each do |key,value|
      cells = key.split(':')
      first_cell = cells[0]
      last_cell = cells[1]
      # need to turn excel cell label like A1 or BB235 into row and column
      # 'A' has codepoint 65
      letters = first_cell[/[A-Za-z]+/].upcase
      cp = letters.codepoints.to_a
      first_col= 0
      if cp.length == 1
        first_col = cp[0] - 64
      elsif cp.length == 2
        first_col = cp[1]-64 + 26*(cp[0]-64)
      end
      first_row = first_cell[/\d+/].to_i

      letters = last_cell[/[A-Za-z]+/].upcase
      cp = letters.codepoints.to_a
      last_col = 0
      if cp.length == 1
        last_col = cp[0] - 64
      elsif cp.length == 2
        last_col = cp[1]-64 + 26*(cp[0]-64)
      end
      last_row = last_cell[/\d+/].to_i

      range = {"first_row" => first_row,"first_col" => first_col, "last_row" => last_row, "last_col" => last_col, "template" => value}
      ranges << range

    end
    
    return ranges
    
  end

  # apply templates to cells to generate RDF
  def create
    la_finder = LaFinder.new
    # loop over confs
    @conf_array.each do |conf|
      # set up output file
      output_name = OUTPUT_DIR + conf["output"]
      output = File.new(output_name,'w')
      cellmap = conf["cellmap"]
      # process the conf entry to set up the cell ranges to be processed
      ranges = setup(cellmap)
      ranges.each do |range|

        # loop over rows
        for row in range["first_row"]..range["last_row"]

        # loop over columns
          for col in range["first_col"]..range["last_col"]
            # set the context for the ERB evaluation
            b = binding
            # read template
            tp = File.read(TEMPLATE_DIR + range["template"])
            # set up the ERB processor
            turtle = ERB.new(tp)
            # process tp
            output << turtle.result(b) << "\r\n\r\n"
          end
        end
      end # of loop over ranges
      output.close
    end # of loop over confs
    
  end





end


# run it

generator = RdfCreator.new("allotments")
generator.create


