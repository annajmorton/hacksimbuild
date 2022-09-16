require 'openstudio'
require 'openstudio/measure/ShowRunnerOutput'
require 'fileutils'
require 'openstudio-standards'
require "rubyXL"
require 'rubyXL/convenience_methods/cell'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL/convenience_methods/worksheet'

#require_relative '../measure.rb'
#require 'minitest/autorun'


translator = OpenStudio::OSVersion::VersionTranslator.new
Proposed_path = OpenStudio::Path.new(File.dirname(__FILE__) + '/proposed_window3.osm')
xlsx_path = File.dirname(__FILE__) + '/output.xlsx'
model = translator.loadModel(Proposed_path)
workbook = RubyXL::Parser.parse(xlsx_path)

model = model.get

worksheet = workbook[0]
worksheet.delete_column(1)
worksheet.delete_column(2)
ws_row = 0 

stories = model.getBuildingStorys

doasnames = ["00", "01 Retail", "02", "LO Office", "UP Office"]
doaszones = [] 

doasnames.each do |dname|
    dname = "DOAS " + dname
    doaszones.push([])
    puts dname
end


stories.each do |story|

    spaces = story.spaces
    n=0
    if story.name.get == "L00"
    
    elsif story.name.get == "L01"
        n=1
    elsif story.name.get == "L02"
        n=2
    elsif story.name.get == "L03 L06" || "L07" || "L08 L11"
        n=3
    else 
        n=4
    end


    spaces.each do |space|
        tz = space.thermalZone.get
        worksheet.add_cell(ws_row, 0, space.name.get) 
        worksheet.add_cell(ws_row, 1, tz.name.get)
        worksheet.add_cell(ws_row, 2, story.name.get)  
        ws_row +=1

        doaszones[n].push(tz)
    end

end 

puts doasnames[3]
doaszones[3].each do |tz|
    puts tz.name.get
end
workbook.write(xlsx_path)