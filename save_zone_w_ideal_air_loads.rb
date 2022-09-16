require 'C:\openstudio-3.2.1\Ruby\openstudio.rb'
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
Proposed_path = OpenStudio::Path.new(File.dirname(__FILE__) + '/Proposed.osm')
xlsx_path = File.dirname(__FILE__) + '/output.xlsx'
model = translator.loadModel(Proposed_path)
workbook = RubyXL::Parser.parse(xlsx_path)

model = model.get

zones = model.getThermalZones
worksheet = workbook["no_HVAC"]
worksheet.delete_column(1)
worksheet.delete_column(2)
ws_row = 0 

zones.each do |zone|
    equipment = zone.equipment
    if equipment.kind_of?(Array)
        if equipment.length == 0
            tzname = zone.name.get
            worksheet.add_cell(ws_row,0, "#{tzname}") 
            puts tzname
            ws_row +=1
        end
    end 
end

workbook.write(xlsx_path)