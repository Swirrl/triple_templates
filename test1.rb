require 'roo'
require 'json'
require 'erb'

class RdfCreator
  SPREADSHEETS_DIR = "spreadsheets/"
  TEMPLATE_DIR = "templates/"
  CONFIG_DIR = "config/"
  OUTPUT_DIR = "outputs/"

  def initialize(project)
    spreadsheet_name = SPREADSHEETS_DIR + project + ".xlsx"
    config_name = CONFIG_DIR + project + ".conf"
    output_name = OUTPUT_DIR + project + ".ttl"

    # open files
    @spreadsheet = Excelx.new(spreadsheet_name)
    config = File.new(config_name,'r')
    @output = File.new(output_name,'w')
    
    # read config
    @conf = JSON.parse(config.read)
    
    # process the config file to set up the cell ranges to be processed
    @ranges = setup(@conf)
    
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
    
    @ranges.each do |range|

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
          @output << turtle.result(b) << "\r\n\r\n"
        end
      end
    end # of loop over ranges
    
  end





end


# run it

generator = RdfCreator.new("housing-starts")
generator.create


