# frozen_string_literal: true

# Require the latest version of OpenStudio (3.2.1)

# require 'C:\openstudio-2.9.1\Ruby\openstudio.rb'
# require 'C:\openstudio-3.1.0\Ruby\openstudio.rb'
require 'C:\openstudio-3.2.1\Ruby\openstudio.rb'
# require '/Applications/OpenStudio-3.2.1/Ruby/openstudio.rb'

require 'openstudio-standards'

# Helper to load a model in one line

def osload(path)
  translator = OpenStudio::OSVersion::VersionTranslator.new

  ospath = OpenStudio::Path.new(path)

  model = translator.loadModel(ospath)

  if model.empty?

    raise "Path '#{path}' is not a valid path to an OpenStudio Model.  Did you remember to use backlashes / instead of forward slashes \ ?"

  else

    model = model.get

  end

  model
end

# Extend ModelObject class to add a to_actual_object method
# Casts a ModelObject into what it actually is (OS:Node for example...)
class OpenStudio::Model::ModelObject
  def to_actual_object
    obj_type = iddObjectType.valueName
    obj_type_name = obj_type.gsub('OS_', '').gsub('_', '')
    method_name = "to_#{obj_type_name}"
    if respond_to?(method_name)
      actual_thing = method(method_name).call
      return actual_thing.get unless actual_thing.empty?
    end
    false
  end
end

def remove_light_sns(model)
  sensors = model.getDaylightingControls

  puts "Found #{sensors.size} daylighting sensors"

  sensors.each do |s|
    puts "Removing #{s.name}"

    s.remove
  end
end

def remove_empty_zones(model)
  model.getThermalZones.each do |z|
    next unless z.spaces.empty?

    unless z.equipment.empty?
      z.equipment.each(&:remove)
    end
    puts "Removing #{z.name}"
    z.remove
  end
end

def remove_empty_spaces(model)
  model.getSpaces.each do |s|
    next unless s.surfaces.empty?

    puts "Removing #{s.name}"
    s.remove
  end
end


def change_airloop(old_airloop, new_airloop)
  tz_v = old_airloop.thermalZones

  tz_v.each do |z|
    # Check if there is some other Airloops attached
    airloop_v = z.airLoopHVACs
    airloop_v.each do |airloop|
      puts "Removing zone #{z.name} from #{airloop.name}"
      airloop.removeBranchForZone(z)
    end
    new_airloop.addBranchForZone(z)
  end
  # Remove the old airloop
  old_airloop.remove
end

def stories_with_airloops(model)
  # Create hash of airloops with their respective stories
  h = {}
  model.getBuildingStorys.each do |s|
    model.getAirLoopHVACs.each do |a|
      next unless a.name.get.include? s.name.get

      h.store(s, a)
    end
  end
  h
end

def airloops_without_stories(model)
  del_airloops = []
  model.getAirLoopHVACs.each do |a|
    # next if a.name.get.include? s.name.get
    next unless a.name.get.include? 'PSZ'

    del_airloops << a
  end
  del_airloops
end

def baseline_vav(model)
  old_airloops = airloops_without_stories(model)
  hash = stories_with_airloops(model)
  old_airloops.each do |o_airloop|
    space = o_airloop.thermalZones[0].spaces[0]
    change_airloop(o_airloop, hash[space.buildingStory.get])
  end
  remove_coils('PSZ-AC', model)
end

def remove_coils(name, model)
  hcc = []
  hcc << model.getCoilCoolingWaters
  hcc << model.getCoilHeatingWaters
  hcc.each do |coils|
    coils.each do |c|
      next unless c.name.get.include? name

      puts c.name
      c.remove
    end
  end
end

# Create a CentralHeatPump System

def central_hp(model)
  hp = OpenStudio::Model::CentralHeatPumpSystem.new(model)
  hp.setName('Multistack MS070XC2H2H2AAC-410A (2 MODULES)')
  ch_module = OpenStudio::Model::CentralHeatPumpSystemModule.new
  ch_module.setNumberofChillerHeaterModules(2)
  perf = OpenStudio::Model::ChillerHeaterPerformanceElectricEIR.new(model)
  c_capacity = OpenStudio.convert(68.37, 'ton', 'W').get
  perf.setReferenceCoolingModeEvaporatorCapacity(c_capacity)
  cooling_COP = 4.83
  perf.setReferenceCoolingModeCOP(cooling_COP)
  c_lwt = OpenStudio.convert(43, 'F', 'C').get
  perf.setReferenceCoolingModeLeavingChilledWaterTemperature(c_lwt)
  con_ewt = OpenStudio.convert(85, 'F', 'C').get
  perf.setReferenceCoolingModeEnteringCondenserFluidTemperature(con_ewt)
  con_lwt = OpenStudio.convert(95, 'F', 'C').get
  perf.setReferenceCoolingModeLeavingCondenserWaterTemperature(con_lwt)
  evap_cap_htg_sim = OpenStudio.convert(900.1 / 12, 'ton', 'W').get
  evap_cap_clg_sim = OpenStudio.convert(55.32, 'ton', 'W').get
  cc_ratio = evap_cap_htg_sim / evap_cap_clg_sim
  perf.setReferenceHeatingModeCoolingCapacityRatio(cc_ratio)
  pw_ratio = 69.21 / 49.74
  perf.setReferenceHeatingModeCoolingPowerInputRatio(pw_ratio)
  h_lchwt = OpenStudio.convert(43, 'F', 'C').get
  perf.setReferenceHeatingModeLeavingChilledWaterTemperature(h_lchwt)
  h_con_lwt = OpenStudio.convert(130, 'F', 'C').get
  perf.setReferenceHeatingModeLeavingCondenserWaterTemperature(h_con_lwt)
  h_con_ewt = OpenStudio.convert(75, 'F', 'C').get
  perf.setReferenceHeatingModeEnteringCondenserFluidTemperature(h_con_ewt)
  h_chwewt = OpenStudio.convert(63, 'F', 'C').get
  perf.setHeatingModeEnteringChilledWaterTemperatureLowLimit(h_chwewt)
  perf.setChilledWaterFlowModeType('VariableFlow')
  chw_f = OpenStudio.convert(82, 'gal/min', 'm^3/s').get
  perf.setDesignChilledWaterFlowRate(chw_f)
  c_f = OpenStudio.convert(198, 'gal/min', 'm^3/s').get
  perf.setDesignCondenserWaterFlowRate(c_f)
  hw_f = OpenStudio.convert(79.55, 'gal/min', 'm^3/s').get
  perf.setDesignHotWaterFlowRate(hw_f)
  perf.setCompressorMotorEfficiency(1)
  ch_module.setChillerHeaterModulesPerformanceComponent(perf)
  hp.addModule(ch_module)
end

def max_humidity_spm(model, airloop, max_hr, min_hr)
  node = airloop.supplyOutletNode
  sp = OpenStudio::Model::SetpointManagerMultiZoneMaximumHumidityAverage.new(model)
  sp.setMaximumSetpointHumidityRatio(min_hr)
  sp.setMinimumSetpointHumidityRatio(max_hr)
  sp.addToNode(node)
end

def dehumidification(m)
  sch = m.getScheduleByName('52F').get
  m.getCoilCoolingWaters.each do |cc|
    node = cc.airOutletModelObject.get.to_Node.get
    spm = OpenStudio::Model::SetpointManagerScheduled.new(m, sch)
    spm.addToNode(node)
    puts node
  end
end

def supply_airflows(model)
  model.getAirLoopHVACs.each do |a|
    node = a.supplyOutletNode
    name = 'System Node Current Density Volume Flow Rate'
    v = OpenStudio::Model::OutputVariable.new(name, model)
    v.setReportingFrequency('Timestep')
    v.setKeyValue(node.name.get)
  end
end

def default_appendixg_curve(model)
  model.getFanVariableVolumes.each do |f|
    f.setFanPowerCoefficient1(0.0013)
    f.setFanPowerCoefficient2(0.1470)
    f.setFanPowerCoefficient3(0.9506)
    f.setFanPowerCoefficient4(-0.0998)
    f.setFanPowerCoefficient5(0)
  end
end

def mult_vav_fc_fans(model)
  model.getFanVariableVolumes.each do |f|
    f.setFanPowerCoefficient1(0.3038)
    f.setFanPowerCoefficient2(-0.7608)
    f.setFanPowerCoefficient3(2.2729)
    f.setFanPowerCoefficient4(-0.8169)
    f.setFanPowerCoefficient5(0)
  end
end


def flat_curve(model)
  model.getFanVariableVolumes.each do |f|
    f.setFanPowerCoefficient1(0.0)
    f.setFanPowerCoefficient2(0.0)
    f.setFanPowerCoefficient3(0.0)
    f.setFanPowerCoefficient4(0.0)
    f.setFanPowerCoefficient5(0)
  end
end

def create_dsoa(model, csv_path)
  require 'csv'
  table = CSV.parse(File.read(csv_path), headers: true)
  table.each do |row|
    next if model.getThermalZoneByName(row['Thermal Zone']).empty?

    tz = model.getThermalZoneByName(row['Thermal Zone']).get
    spaces = tz.spaces
    spaces.each do |s| 
      next if s.designSpecificationOutdoorAir.empty?

      s.designSpecificationOutdoorAir.get.remove
    end
  end

  table.each do |row|
    next if model.getThermalZoneByName(row['Thermal Zone']).empty?

    tz = model.getThermalZoneByName(row['Thermal Zone']).get
    puts "Matched zone #{tz.name}"
    s_v = tz.spaces
    tz_area = tz.floorArea
    tz_ppl = tz.numberOfPeople
    a_flow = row['ABSOLUTE MIN CFM'].to_f
    ppl_flow = row['MIN (CFM)'].to_f - a_flow
    # Check if a DSOA does not exist in the first space of the zone
    if s_v[0].designSpecificationOutdoorAir.empty?
      dsoa = OpenStudio::Model::DesignSpecificationOutdoorAir.new(model)
    else
      dsoa = s_v[0].designSpecificationOutdoorAir.get
      o_a_flow_si = dsoa.outdoorAirFlowperFloorArea * tz_area
      o_a_flow = OpenStudio.convert(o_a_flow_si, 'm^3/s', 'cfm').get
      a_flow += o_a_flow
      o_ppl_flow_si = dsoa.outdoorAirFlowperPerson * tz_ppl
      o_ppl_flow = OpenStudio.convert(o_ppl_flow_si, 'm^3/s', 'cfm').get
      ppl_flow += o_ppl_flow
    end
    dsoa.setName("#{tz.name} DSOA")
    si_a_flow = OpenStudio.convert(a_flow / tz_area, 'cfm', 'm^3/s').get
    dsoa.setOutdoorAirFlowperFloorArea(si_a_flow)
    unless tz_ppl.zero?
      si_ppl_flow = OpenStudio.convert(ppl_flow / tz_ppl, 'cfm', 'm^3/s').get
      dsoa.setOutdoorAirFlowperPerson(si_ppl_flow)
      # ppl = s_v[0].people[0]
      # ppl_sch = ppl.numberofPeopleSchedule.get
      # ppl_sch = model.getScheduleByName('Medium Office Bldg Occ').get
      # dsoa.setOutdoorAirFlowRateFractionSchedule(ppl_sch)
    end
    s_v.each { |s| s.setDesignSpecificationOutdoorAir(dsoa) }
  end
end


# def adjust_vav_terminals(model,run_dir)
#   sql_path = OpenStudio::Path.new(run_dir)
#   sql = std.safe_load_sql(sql_path)
#   model.setSqlFile(sql)
#   model.getAirLoopHVACs.each do |a|
#     std.air_loop_hvac_apply_minimum_vav_damper_positions(a)
#     std.air_loop_hvac_adjust_minimum_vav_damper_positions(a)
#   end
# end

# model.getScheduleConstants.each do |s|
#   next unless s.name.get.include? 'Measure'

#   s.remove
# # end



def vsd_dp_reset_pumps(model)
  model.getPumpVariableSpeeds.each do |pump|
    pump.setCoefficient1ofthePartLoadPerformanceCurve(0)
    pump.setCoefficient2ofthePartLoadPerformanceCurve(0.0205)
    pump.setCoefficient3ofthePartLoadPerformanceCurve(0.4101)
    pump.setCoefficient4ofthePartLoadPerformanceCurve(0.5753)
  end
end

def remove_hstats(model)
  model.getThermalZones.each do |tz|
    next unless tz.airLoopHVACs.empty?

    next if tz.zoneControlHumidistat.empty?

    puts "removing from #{tz.name}"
    tz.zoneControlHumidistat.get.remove
  end
end

def report_oa_airloop(airloop)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  airloop_oa_si = 0
  airloop.thermalZones.each do |tz|
    next unless std.thermal_zone_outdoor_airflow_rate(tz) > 0
    airflow_oa = std.thermal_zone_outdoor_airflow_rate(tz) * tz.multiplier
    airloop_oa_si += airflow_oa
    puts "#{tz.name} has #{OpenStudio.convert(airflow_oa, 'm^3/s', 'cfm')} cfm of OA" 
  end
  airloop_oa = OpenStudio.convert(airloop_oa_si, 'm^3/s', 'cfm').get
  puts "Airloop #{airloop.name} Outdoor Airflow is  #{airloop_oa.round(0)} cfm"
  return airloop_oa.round(0)
end

def report_oa(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  bld_oa_si = 0
  model.getAirLoopHVACs.sort.each do |hvac|
    airloop_oa = report_oa_airloop(hvac)
    puts "Airloop #{hvac.name} requirement is #{airloop_oa} cfm"
  end
  model.getThermalZones.each do |tz|
    bld_oa_si += std.thermal_zone_outdoor_airflow_rate(tz) * tz.multiplier
  end
  bld_oa = OpenStudio.convert(bld_oa_si, 'm^3/s', 'cfm')
  puts "Building OA requirement is #{bld_oa} cfm"
end

def report_oa_per_story(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  model.getBuildingStorys.each do |st|
    story_oa_si = 0
    area = 0.0
    st.spaces.each do |s|
      if s.thermalZone.empty?
        puts "Space #{s.name} does not have a TZ, skipping"
        next
      end
      tz = s.thermalZone.get
      story_oa_si += std.thermal_zone_outdoor_airflow_rate(tz) * tz.multiplier
      area = area + (s.floorArea)
    end
    bld_oa = OpenStudio.convert(story_oa_si, 'm^3/s', 'cfm')
    puts "Story #{st.name} requirement is #{bld_oa} cfm"
  end
end

def design_oa_fcu(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  model.getZoneHVACFourPipeFanCoils.each do |fcu|

    tz = fcu.thermalZone.get
    fcu.setMaximumOutdoorAirFlowRate(std.thermal_zone_outdoor_airflow_rate(tz))
    puts "#{fcu.name} requirement is #{OpenStudio.convert(fcu.maximumOutdoorAirFlowRate.get, 'm^3/s', 'cfm')} cfm"

  end
end

def hvac_zones(model)
  model.getAirLoopHVACs.each do |hvac|
    puts hvac.name
    hvac.thermalZones.each { |z| puts z.name }
  end
end

def delete_interior_walls(model)
  vector = []
  model.getSurfaces.each do |s|
    next unless s.outsideBoundaryCondition == 'Surface'

    vector << s
  end
  model.getSurfaces.each do |s|
    next unless s.outsideBoundaryCondition == 'Adiabatic'

    vector << s
  end

  vector.each(&:remove)
end

def delete_roofs(model)
  model.getSurfaces.each do |s|
    next unless s.surfaceType == 'RoofCeiling'

    s.remove
  end
end

def del_uh_wv(model)
  model.getZoneHVACUnitHeaters.each do |uh|
    next unless uh.name.get.to_s.include? 'WC'

    puts "Removing #{uh.name}"
    uh.remove
  end
end

def unit_spaces_areas(model, spacetype = true)
  st = spacetype
  units = model.getBuildingUnits.sort
  units.each do |bu|
    puts "#{bu.name}, #{OpenStudio.convert(bu.floorArea, 'm^2', 'ft^2').get} ft2"
    puts '___________'
    bu.spaces.sort.each { |s| puts "#{s.name} has #{OpenStudio.convert(s.floorArea, 'm^2', 'ft^2').get} ft2" }
    puts '___________'
  end
end

def reassign_vrf(unit_name,model)
  cu = model.getAirConditionerVariableRefrigerantFlowByName(unit_name).get

  model.getAirConditionerVariableRefrigerantFlows.each do |vrf|
    next if vrf == cu

    vrf.terminals.each do |t|
      nt = t.clone
      nt = nt.to_actual_object
      nt.addToThermalZone(t.thermalZone.get)
      cu.addTerminal(nt)
      puts "added #{nt.name}"
    end
    vrf.removeAllTerminals
  end
end

def attach_sql(model, sql_path)
  sql_file = OpenStudio::SqlFile.new(OpenStudio::Path.new(sql_path))
  model.setSqlFile(sql_file)
  model
end

def close_sql(model)
  sql_file = model.sqlFile.get
  sql_file.close
end

def report_baseline_vavs(model, sql_path)
  attach_sql(model, sql_path)
  model.getAirLoopHVACs.each do |airloop|
    vav_supply_flow(model, airloop)
    vav_clg_cap(model,airloop)
    vav_htg_cap(model, airloop)
    report_oa_airloop(airloop)
  end
  # model.sqlFile.get.close
  close_sql(model)
end

def vav_supply_flow(model, airloop)
  tot_airflow = 0
  components = airloop.supplyComponents
  components.each do |sc|
    next if sc.to_FanVariableVolume.empty? && sc.to_FanConstantVolume.empty?

    sc = sc.to_actual_object


    fan_flow = OpenStudio.convert(model.getAutosizedValue(sc, 'Design Size Maximum Flow Rate', 'm3/s').get, 'm^3/s', 'cfm').get


    # puts "#{coil.name} has a #{coil_cap}"
    tot_airflow += fan_flow
  end
  puts "Airloop #{airloop.name} supply airflow is  #{tot_airflow.round(0)} cfm"
  tot_airflow
end

def report_baseline_FCUs(model,sql_path)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  attach_sql(model, sql_path)
  design_airflow = 0
  c_capacity = 0
  h_capacity = 0
  fcu_oa_si = 0
  model.getZoneHVACFourPipeFanCoils.each do |fcu|
    design_airflow += OpenStudio.convert(model.getAutosizedValue(fcu, 'Design Size Maximum Supply Air Flow Rate', 'm3/s').get, 'm^3/s', 'cfm').get
    # vav_clg_cap(model,airloop)
    h_coil = fcu.heatingCoil
    c_coil = fcu.heatingCoil
    h_capacity += OpenStudio.convert(model.getAutosizedValue(h_coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    c_capacity += OpenStudio.convert(model.getAutosizedValue(c_coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    tz = fcu.thermalZone.get
    fcu_oa_si += std.thermal_zone_outdoor_airflow_rate(tz)
  end
  fcu_oa = OpenStudio.convert(fcu_oa_si, 'm^3/s', 'cfm')
  puts "The building Design Airflow for the 4pipeFCUs is  #{design_airflow} cfm"
  puts "The building heating capacity for the 4pipeFCUs is  #{capacity} BTU/h"
  puts "The building Outdoor Airflow is  #{fcu_oa} cfm"
  close_sql(model)
end

def report_baseline_PTACs(model,sql_path)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  attach_sql(model, sql_path)
  sqlFile = model.sqlFile.get
  design_airflow = 0
  c_capacity = 0
  h_capacity = 0
  fcu_oa_si = 0
  fan_power = 0
  model.getZoneHVACPackagedTerminalAirConditioners.each do |fcu|
    design_airflow += OpenStudio.convert(model.getAutosizedValue(fcu, 'Design Size Cooling Supply Air Flow Rate', 'm3/s').get, 'm^3/s', 'cfm').get
    # vav_clg_cap(model,airloop)
    h_coil = fcu.heatingCoil
    c_coil = fcu.coolingCoil
    fan = fcu.supplyAirFan
    h_capacity += OpenStudio.convert(model.getAutosizedValue(h_coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    c_capacity += OpenStudio.convert(model.getAutosizedValue(c_coil, 'Design Size Gross Rated Total Cooling Capacity', 'W').get, 'W', 'Btu/h').get
    tz = fcu.thermalZone.get
    query = "SELECT Value FROM tabulardatawithstrings WHERE ReportName='EquipmentSummary' and TableName='Fans' and RowName= '#{fan.name.get.upcase}' and ColumnName= 'Rated Electricity Rate' and ReportForString= 'Entire Facility'"
    results = sqlFile.execAndReturnFirstString(query).get
    fan_power += results.to_f
    fcu_oa_si += std.thermal_zone_outdoor_airflow_rate(tz)
  end
  fcu_oa = OpenStudio.convert(fcu_oa_si, 'm^3/s', 'cfm')
  puts "The building Design Airflow for the PTACs is  #{design_airflow} cfm"
  puts "The building cooling capacity for the PTACs is  #{c_capacity} BTU/h"
  puts "The building heating capacity for the PTACs is  #{h_capacity} BTU/h"
  puts "The building fan power for the PTACs is  #{fan_power} W"
  puts "The building Outdoor Airflow is  #{fcu_oa} cfm"
  close_sql(model)
end

def report_chillers_and_boiler_capacities(model,sql_path)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  attach_sql(model, sql_path)
  design_airflow = 0
  c_capacity = 0
  h_capacity = 0
  fcu_oa_si = 0
  model.getZoneHVACFourPipeFanCoils.each do |fcu|
    design_airflow += OpenStudio.convert(model.getAutosizedValue(fcu, 'Design Size Maximum Supply Air Flow Rate', 'm3/s').get, 'm^3/s', 'cfm').get
    # vav_clg_cap(model,airloop)
    h_coil = fcu.heatingCoil
    c_coil = fcu.heatingCoil
    h_capacity += OpenStudio.convert(model.getAutosizedValue(h_coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    c_capacity += OpenStudio.convert(model.getAutosizedValue(c_coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    tz = fcu.thermalZone.get
    fcu_oa_si += std.thermal_zone_outdoor_airflow_rate(tz)
  end
  fcu_oa = OpenStudio.convert(fcu_oa_si, 'm^3/s', 'cfm')
  puts "The building Design Airflow for the 4pipeFCUs is  #{design_airflow} cfm"
  puts "The building heating capacity for the 4pipeFCUs is  #{capacity} BTU/h"
  puts "The building Outdoor Airflow is  #{fcu_oa} cfm"
  close_sql(model)
end


def vav_clg_cap(model, airloop)
  tot_capacity = 0
  components = airloop.supplyComponents
  components.each do |sc|
    next if sc.to_CoilCoolingWater.empty? && sc.to_CoilCoolingDXSingleSpeed.empty?

    sc = sc.to_actual_object

    unless sc.to_CoilCoolingWater.empty?
      coil_cap = OpenStudio.convert(model.getAutosizedValue(sc, 'Design Size Design Coil Load', 'W').get, 'W', 'Btu/h').get
    end

    unless sc.to_CoilCoolingDXSingleSpeed.empty?
      coil_cap = OpenStudio.convert(model.getAutosizedValue(sc, 'Design Size Gross Rated Total Cooling Capacity', 'W').get, 'W', 'Btu/h').get
    end
    # puts "#{coil.name} has a #{coil_cap}"
    tot_capacity += coil_cap
  end
  puts "Airloop #{airloop.name} Cooling capacity is is  #{tot_capacity.round(0)} BTUH"
  tot_capacity
end

def vav_areas(model, airloop)
  tot_areas = 0
  zones = airloop.thermalZones
  zones.each do |zone|
    zone_floor_area_si = zone.floorArea * zone.multiplier
    zone_floor_area_ip = OpenStudio.convert(zone_floor_area_si, 'm^2', 'ft^2').get
    tot_areas += zone_floor_area_ip
  end
  puts "Airloop #{airloop.name} served area is #{tot_areas.round(0)} ft2"
  tot_areas
end

def vav_htg_cap(model, airloop)
  tot_capacity = 0
  zones = airloop.thermalZones
  capacity = 0
  components = airloop.supplyComponents
  components.each do |sc|
    next if sc.to_CoilHeatingWater.empty? && sc.to_CoilHeatingGas.empty? 

    sc = sc.to_actual_object

    unless sc.to_CoilHeatingWater.empty?
      coil_cap = OpenStudio.convert(model.getAutosizedValue(sc, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    end

    unless sc.to_CoilHeatingGas.empty?
      coil_cap = OpenStudio.convert(model.getAutosizedValue(sc, 'Design Size Nominal Capacity', 'W').get, 'W', 'Btu/h').get
    end
    # puts "#{coil.name} has a #{coil_cap}"
    capacity += coil_cap
  end
  equipment = []
  zones.each { |z| equipment << z.equipment }
  equipment.flatten!
  equipment.each do |e|
    next if e.to_AirTerminalSingleDuctVAVReheat.empty?

    e = e.to_actual_object
    coil = e.reheatCoil.to_actual_object
    coil_cap = OpenStudio.convert(model.getAutosizedValue(coil, 'Design Size Rated Capacity', 'W').get, 'W', 'Btu/h').get
    # puts "#{coil.name} has a #{coil_cap}"
    capacity += coil_cap
  end
  puts "Airloop #{airloop.name} heating capacity is #{capacity.round(0)} BTU/h"
  capacity
end

def replace_windows(seed_model, target_model)
  target_model.getSubSurfaces.each do |s|
    name = s.name.get
    s_s = seed_model.getSubSurfaceByName(name).get
    v_s = s_s.vertices
    s.setVertices(v_s)
  end
end

def replace_windows_2(seed_model, target_model)
  target_model.getSubSurfaces.each(&:remove)
  seed_model.getSubSurfaces.each do |s|
    surface_name = s.surface.get.name.get
    t_s = target_model.getSurfaceByName(surface_name).get
    n_s_s = s.clone(target_model)
    n_s_s = n_s_s.to_actual_object
    n_s_s.setSurface(t_s)
  end
end

def print_surfaces(space)
  sp = model.getSpaceByName(space).get
  space.surfaces.each do |s|
    puts "#{s.name} is a #{s.surfaceType}" 
  end
end

def replace_terminals(airloop_name,model)
  airloop = model.getAirLoopHVACByName(airloop_name).get
  sch = model.getScheduleByName('Always On Discrete').get
  tzs = airloop.thermalZones
  tzs.each do |z| 
    airloop.removeBranchForZone(z)
    t = OpenStudio::Model::AirTerminalSingleDuctVAVNoReheat.new(model,sch)
    t.setZoneMinimumAirFlowInputMethod('Fixed')
    t.setFixedMinimumAirFlowRate(0)
    t.setControlForOutdoorAir(true)
    airloop.addBranchForZone(z,t)
  end
end

def copy_dsoa(seed_model, target_model)
  seed_model.getSpaces.each do |s|
    next if s.designSpecificationOutdoorAir.empty?
    
    dsoa = s.designSpecificationOutdoorAir.get
    s_t = target_model.getSpaceByName(s.name.get).get
    dsoa_t = dsoa.clone(target_model)
    dsoa_t = dsoa_t.to_actual_object
    s_t.setDesignSpecificationOutdoorAir(dsoa_t)
  end
  report_oa(seed_model)
  report_oa(target_model) 
end

def baseline_vav_oa(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  bld_oa_si = 0
  # Change the OA in the autosizing
  model.getAirLoopHVACs.each do |hvac|
    airloop_oa_si = 0
    hvac.thermalZones.each do |tz|
      airloop_oa_si += std.thermal_zone_outdoor_airflow_rate(tz)
    end
    sizing = hvac.sizingSystem
    sizing.setDesignOutdoorAirFlowRate(airloop_oa_si)
    airloop_oa = OpenStudio.convert(airloop_oa_si, 'm^3/s', 'cfm')
    puts "Airloop #{hvac.name} requirement is #{airloop_oa} cfm"
  end

  # Change the minimum airflow for each 
  # Control for OA

end

def change_availability_manager(model)

  model.getAvailabilityManagerNightCycles.each do |a|
    a.setCyclingRunTimeControlType('Thermostat')
    a.setThermostatTolerance(0.277777777777778) 
  end

end

def reduce_static(model)

  model.getFanVariableVolumes.each do |f|
    puts f.name
    new_eff = f.fanTotalEfficiency*0.9
    f.setFanTotalEfficiency(new_eff)
  end
end

def eliminate_erv_power(model)
  model.getHeatExchangerAirToAirSensibleAndLatents.each do |hx|
    hx.setNominalElectricPower(0.0)
  end
end

def leed_spacetypes(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  spacetypes = model.getSpaceTypes.sort
  header = ['Spacetype', 'Area (ft^2)', 'LPD (W/ft^2)', 'EPD (W/ft^2)', 'Outdoor Air (cfm)']
  spacetype_attributes_vector = []
  spacetypes.each do |st|
    spacetype_attributes = []
    spacetype_attributes << st.name
    spacetype_attributes << OpenStudio.convert(st.floorArea, 'm^2', 'ft^2').get.round(0)
    lpd_and_epd = enhancedLightingAndEquipmentPowerPerFloorArea(st)
    spacetype_attributes << lpd_and_epd[0]
    spacetype_attributes << lpd_and_epd[1]
    tzs = []
    st.spaces.each do |s|
      next if s.thermalZone.empty?
      tzs << s.thermalZone.get
    end
    tzs = tzs.uniq
    ventilation_cfm = 0
    tzs.each do |z|
      ventilation_cfm = ventilation_cfm + (OpenStudio.convert(std.thermal_zone_outdoor_airflow_rate(z), 'm^3/s', 'cfm').get.round(0) * z.multiplier)
    end
    spacetype_attributes << ventilation_cfm
    spacetype_attributes_vector << spacetype_attributes
  end
  File.open('leed_spacetypes_report.csv', 'w') do |file|
    file.puts header.join(',')
    spacetype_attributes_vector.each do |row|
      file.puts row.join(',')
    end
  end
end

def energy_star_spaces(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  spacetypes = model.getSpaceTypes.sort
  spacetypes.each do |st|
    puts "#{st.name}, #{OpenStudio.convert(st.floorArea, 'm^2', 'ft^2').get.round(0)} ft2"
    puts '___________'
    enhancedLightingAndEquipmentPowerPerFloorArea(st)
    tzs = []
    st.spaces.each do |s|
      next if s.thermalZone.empty?
      tzs << s.thermalZone.get
    end
    tzs = tzs.uniq
    ventilation_cfm = 0
    tzs.each do |z|
      ventilation_cfm = ventilation_cfm + OpenStudio.convert(std.thermal_zone_outdoor_airflow_rate(z), 'm^3/s', 'cfm').get.round(0) * z.multiplier
    end
    puts "#{st.name}, outdoor air is  #{ventilation_cfm} cfm"
    puts '___________'
  end
end

def ventilation_thermalzone_table(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  cfm = 0.0
  header =["Space", "Ventilation (cfm)"]
  space = []

  model.getThermalZones.sort.each do |z|
    spaces_items = []
    ventilation_cfm = OpenStudio.convert(std.thermal_zone_outdoor_airflow_rate(z), 'm^3/s', 'cfm').get * z.multiplier
    # next if ventilation_cfm == 0.0
    cfm += ventilation_cfm
    spaces_items << z.name
    spaces_items << ventilation_cfm
    space << spaces_items 
    puts "#{z.name}, outdoor air is  #{ventilation_cfm.round(0)} cfm, zone multiplier #{z.multiplier}"
  end
  puts "Building is  #{cfm.round(0)} cfm"

  File.open('report.csv', 'w') do |file|
    file.puts header.join(',')
    space.each do |row|
      file.puts row.join(',')
    end
  end
end


def enhancedLightingAndEquipmentPowerPerFloorArea(st)
  results = []
  lp = 0
  ep = 0
  st.spaces.each do |s|
    next if s.thermalZone.empty?
  
    multiplier = s.thermalZone.get.multiplier
    lp = lp + (s.lightingPower * multiplier)
    ep = ep + (s.electricEquipmentPower * multiplier)
  end
  si_st_area = st.floorArea
  results << OpenStudio.convert(lp/si_st_area, 'W/m^2', 'W/ft^2').get.round(2)
  results << OpenStudio.convert(ep/si_st_area, 'W/m^2', 'W/ft^2').get.round(2)
  results
end


def building_lpd(model)
  lp = 0
  ep = 0
  si_bldg_area = 0
  model.getSpaces.each do |s|
    next if s.thermalZone.empty?
  
    multiplier = s.thermalZone.get.multiplier
    lp = lp + (s.lightingPower * multiplier)
    ep = ep + (s.electricEquipmentPower * multiplier)
    si_bldg_area = si_bldg_area + (s.floorArea * multiplier)
  end
  puts "The building has #{lp.round(2)} W for lighting"
  puts '___________'
  puts "The building has #{ep.round(2)} W for equipment"
  puts '___________'
end
  


def leed_airloops(model,sql_path)
  # Create the header for the table
  header =["Airloop","Area", "Total Cooling Capacity", "Total Heating Capacity", "Supply Airflow", "Outdoor Airflow"]
  # Get all the model Condensing Units
  airloops_array = []
  attach_sql(model, sql_path)
  model.getAirLoopHVACs.sort.each do |airloop|
    a_hvac_items = []
    a_hvac_items << airloop.name
    a_hvac_items << vav_areas(model, airloop)
    a_hvac_items << vav_clg_cap(model, airloop)
    a_hvac_items << vav_htg_cap(model, airloop)
    a_hvac_items << vav_supply_flow(model, airloop)
    a_hvac_items << report_oa_airloop(airloop)
    airloops_array << a_hvac_items  
  end 
  # if we want this report could write out a csv, html, or any other file here
  puts "Writing CSV report 'leed_airloops_report.csv'"
  File.open('leed_airloops_report.csv', 'w') do |file|
    file.puts header.join(',')
    airloops_array.each do |row|
      file.puts row.join(',')
    end
  end
  close_sql(model)
end

def leed_ptacs(model,sql_path)
  # Create the header for the table
  header =["PTAC","Area", "Total Cooling Capacity", "Total Heating Capacity", "Supply Airflow", "Outdoor Airflow"]
  # Get all the model Condensing Units
  airloops_array = []
  attach_sql(model, sql_path)
  model.getAirLoopHVACs.sort.each do |airloop|
    a_hvac_items = []
    a_hvac_items << airloop.name
    a_hvac_items << vav_areas(model, airloop)
    a_hvac_items << vav_clg_cap(model, airloop)
    a_hvac_items << vav_htg_cap(model, airloop)
    a_hvac_items << vav_supply_flow(model, airloop)
    a_hvac_items << report_oa_airloop(airloop)
    airloops_array << a_hvac_items  
  end 
  # if we want this report could write out a csv, html, or any other file here
  puts "Writing CSV report 'leed_airloops_report.csv'"
  File.open('leed_airloops_report.csv', 'w') do |file|
    file.puts header.join(',')
    airloops_array.each do |row|
      file.puts row.join(',')
    end
  end
  close_sql(model)
end


def zone_sizing(cold_supply, hot_supply)

end


def correct_sizing(model, vrf_name)

end

def put_units_thermalzones(model)
  model.getBuildingUnits.each do |bu|

  end
end

def put_space_type_tz_name(model)
  spacetypes = model.getSpaceTypes.sort
  spacetypes.each do |st|
    puts "#{st.name}, #{OpenStudio.convert(st.floorArea, 'm^2', 'ft^2').get} ft2"
    puts '___________'
    st.spaces.sort.each do |s|
       next if s.thermalZone.empty?

      puts "#{s.thermalZone.get.name}"
    end
    puts '___________'
  end
end

def assign_thermostats(clg_sch,htg_sch,model)
  model.getThermalZones.each do |z|
    next if z.equipment.size == 0

    if z.thermostat.empty?
      thermostat = OpenStudio::Model::ThermostatSetpointDualSetpoint.new(model)
      z.setThermostatSetpointDualSetpoint(thermostat)
    else
      thermostat = z.thermostatSetpointDualSetpoint.get
    end
    thermostat.setCoolingSchedule(clg_sch)
    thermostat.setHeatingSchedule(htg_sch)
  end
end

def assign_thermostats_per_st(clg_sch, htg_sch, st, model)
  model.getThermalZones.each do |z|
    s = z.spaces[0]
    next if s.spaceType.empty?

    next unless s.spaceType.get == st

    if z.thermostat.empty?
      thermostat = OpenStudio::Model::ThermostatSetpointDualSetpoint.new(model)
      z.setThermostatSetpointDualSetpoint(thermostat)
    else
      thermostat = z.thermostatSetpointDualSetpoint.get
    end
    thermostat.setCoolingSchedule(clg_sch)
    thermostat.setHeatingSchedule(htg_sch)
  end
end

def assign_humidistats_per_st(dehum_sch, hum_sch, st, model)
  model.getThermalZones.each do |z|
    s = z.spaces[0]
    next if s.spaceType.empty?

    next unless s.spaceType.get == st

    if z.zoneControlHumidistat.empty?
      humidistat = OpenStudio::Model::ZoneControlHumidistat.new(model)
      z.setZoneControlHumidistat(humidistat)
    else
      humidistat = z.zoneControlHumidistat.get
    end
    humidistat.setDehumidifyingRelativeHumiditySetpointSchedule(dehum_sch)
    humidistat.setHumidifyingRelativeHumiditySetpointSchedule(hum_sch)
  end
end

def use_ideal_loads(model)
  model.getThermalZones.each do |z|
    next if z.thermostat.empty?

    z.setUseIdealAirLoads(true)
  end
end

# original = model.getModelObjectByName('Lights Definition 1').get.to_actual_object

# table.each do |row|
#   newone = original.clone
#   newone = newone.to_actual_object
#   name = row['type']
#   watts = row['Watts / Lamp'].to_f
#   newone.setName(name)
#   newone.setLightingLevel(watts)

# end


def baseline_fanpower(model,sql_path,csv_path)
  attach_sql(model, sql_path)
  table = CSV.parse(File.read(csv_path), headers: true)
  table.each do |row|
    airloop_opt = model.getAirLoopHVACByName(row['airloop'])
    if airloop_opt.empty?
        puts "We could not find #{row['airloop']}in the model"
      next
    end
    airloop = airloop_opt.get
    fan = airloop.supplyFan.get.to_actual_object
    flowrate = fan.autosizedMaximumFlowRate.get
    pressure = fan.pressureRise
    power = row['power (W)'].to_f
    efficiency = pressure * flowrate / power
    fan.setFanEfficiency(efficiency)
    puts "Fan efficiency for #{airloop.name} is #{efficiency.round(3)*100}%"
  end
  close_sql(model)
end

def rename_zones(model)
  model.getThermalZones.each do |z|
    n = z.name.get 
    n.slice!("Thermal ")
    z.setName(n)
  end
end

# std.model_add_chw_loop(model,
#   system_name: 'Chilled Water Loop',
#   cooling_fuel: 'DistrictCooling',
#   dsgn_sup_wtr_temp: 44.0,
#   dsgn_sup_wtr_temp_delt: 10.1,
#   chw_pumping_type: 'const_pri',
#   chiller_cooling_type: nil,
#   chiller_condenser_type: nil,
#   chiller_compressor_type: nil,
#   num_chillers: 1,
#   condenser_water_loop: nil,
#   waterside_economizer: 'none')


# zones = []
# model.getThermalZones.each do |z|
#   next if z.thermostat.empty?

#   zones << z

# end

# std.model_add_four_pipe_fan_coil(model,
#   zones,
#   chilled_water_loop,
#   hot_water_loop: hot_water_loop,
#   ventilation: true,
#   capacity_control_method: 'CyclingFan')

# model.getThermalZones.each do |z|
#   next unless z.equipment.empty?

#   ventilation = false

#   puts z.name

#   z.spaces.each do |s|
#     next if s.designSpecificationOutdoorAir.empty?

#     puts s.designSpecificationOutdoorAir.get

#     ventilation = true
#   end

#   next unless ventilation

#   puts z.name
# end

# model.getFanOnOffs.each do |f|
#   f.setFanEfficiency(0.4267)
# end

# space_types = []
# space_types << 'Fitness Center'
# space_types << 'HighriseApartment Apartment'
# space_types << 'HighriseApartment Office'
# space_types << 'LargeHotel Banquet'
# space_types << 'LargeHotel Corridor'
# space_types << 'LargeHotel Lobby'


# space_types.each do |st_n|
#   st = model.getSpaceTypeByName(st_n)
#   next if st.empty?
#   puts "Using #{st_n}"
#   st = st.get
#   assign_thermostats_per_st(clg_sch, htg_sch, st, model)
# end

# seed_model.getSpaceTypes.each do |st|
#   st.clone(target_model)
# end

def delete_empty_spacetypes(model)
  model.getSpaceTypes.each do |st|
    next unless st.spaces.empty?

    puts "Removing #{st.name}"
    st.remove
  end
end

def delete_empty_stories(model)
  model.getBuildingStorys.each do |st|
    next unless st.spaces.empty?

    puts "Removing #{st.name}"
    st.remove
  end
end

# surfaces = []
# surfaces << 'SURFACE 1868'
# surfaces << 'SURFACE 22'
# surfaces << 'SURFACE 137'
# surfaces << 'SURFACE 1633'
# surfaces << 'SURFACE 1759'
# surfaces << 'SURFACE 1640'
# surfaces << 'SURFACE 718'
# surfaces << 'SURFACE 586'
# surfaces << 'SURFACE 770'

# construction_name = 'Interior Ceiling'

# construction = model.getModelObjectByName(construction_name)
# construction = construction.get
# construction = construction.to_actual_object

# surfaces.each do |s|

#   s = model.getSurfaceByName(s).get

#   puts s.name
#   s.setOutsideBoundaryCondition('Adiabatic')
#   s.setConstruction(construction)
# end

# model.getZoneHVACTerminalUnitVariableRefrigerantFlows.each do |vrf|
  # n = vrf.name.get
  # vrf.supplyAirFan.setName("#{n} Fan")
  # vrf.coolingCoil.get.setName("#{n} CLG COIL")
  # vrf.heatingCoil.get.setName("#{n} HTG COIL")
# end

def fix_vrf_terminal_units(model)
  if model.getScheduleByName('Always off - Measure').empty?
    off_sch = OpenStudio::Model::ScheduleConstant.new(model)
    off_sch.setName('Always off - Measure')
    off_sch.setValue(0)
  else
    off_sch = model.getScheduleByName('Always off - Measure').get
  end
  model.getZoneHVACTerminalUnitVariableRefrigerantFlows.each do |vrf|
    n = vrf.name.get
    fan = vrf.supplyAirFan.to_actual_object
    fan.setName("#{n} Fan")
    c_coil = vrf.coolingCoil.get 
    c_coil.setName("#{n} CLG COIL")
    h_coil = vrf.heatingCoil.get 
    h_coil.setName("#{n} HTG COIL")
    fan.setFanTotalEfficiency(0.7)
    # This is normal
    fan.setMotorEfficiency(0.9)
    fan.setPressureRise(0.0)
    fan.setEndUseSubcategory('VRF fans')
    vrf.setOutdoorAirFlowRateDuringCoolingOperation(0.0)
    vrf.setOutdoorAirFlowRateDuringHeatingOperation(0.0)
    vrf.setOutdoorAirFlowRateWhenNoCoolingorHeatingisNeeded(0.0)
    vrf.setSupplyAirFlowRateWhenNoCoolingisNeeded(0.0)
    vrf.setSupplyAirFlowRateWhenNoHeatingisNeeded(0.0)
    # TODO: expand schedule to add parasitic energy use
    vrf.setZoneTerminalUnitOffParasiticElectricEnergyUse(0.0)
    vrf.setZoneTerminalUnitOnParasiticElectricEnergyUse(0.0)
    vrf.setSupplyAirFanOperatingModeSchedule(off_sch)
  end
end

def area_per_story(model)
  model.getBuildingStorys.each do |st|
    
    area = 0.0
    st.spaces.each do |s|
      area = area + (s.floorArea)
    end
    puts "#{st.name} has #{OpenStudio.convert(area, 'm^2', 'ft^2').get} ft2"
  end
end

def lighting_area_per_space_for_story(model)
  header = ['Thermal Zone', 'Light Plan', 'Multiplier', 'Space Type', 'Area SF']
  table = []
  model.getBuildingStorys.each do |st|
    st.spaces.each do |s|
      values = []
      tz = s.thermalZone.get
      values << tz.name
      values << 'X'
      values << tz.multiplier
      values << s.spaceType.get.name
      values << OpenStudio.convert(s.floorArea*tz.multiplier, 'm^2', 'ft^2').get.round(1)
      table << values
    end
  end
  File.open('lighting_area_per_space_for_story_report.csv', 'w') do |file|
    file.puts header.join(',')
    table.each do |row|
      file.puts row.join(',')
    end
  end
end  

def export_ventilation_areas_per_TZ(model)
  header =["Space", "Space Type", "Area", "Number of People"]
  space = []
  model.getBuildingStorys.each do |st|
    st.spaces.each do |s|
      z_m = s.thermalZone.get.multiplier
      spaces_items = []
      spaces_items << s.name.get
      spaces_items << s.spaceType.get.name
      spaces_items << OpenStudio.convert(s.floorArea*z_m, 'm^2', 'ft^2').get.round(1)
      spaces_items << s.numberOfPeople.round(0)
      space << spaces_items  
    end
  end

  File.open('report.csv', 'w') do |file|
    file.puts header.join(',')
    space.each do |row|
      file.puts row.join(',')
    end
  end
end

def export_ventilation_areas_per_TZ(model)
  header =["Thermal Zone", "Space Type", "Area", "Number of People"]
  space = []
  model.getThermalZones.each do |tz|
    z_m = tz.multiplier
    spaces_items = []
    spaces_items << tz.name.get
    spaces_items << tz.spaces[0].spaceType.get.name
    spaces_items << OpenStudio.convert(tz.floorArea*z_m, 'm^2', 'ft^2').get.round(1)
    spaces_items << tz.numberOfPeople.round(0)
    space << spaces_items  
  end

  File.open('report.csv', 'w') do |file|
    file.puts header.join(',')
    space.each do |row|
      file.puts row.join(',')
    end
  end
end


def remove_orphan_vrf_terminals(model)
  units = []
  model.getAirConditionerVariableRefrigerantFlows.each do |vrf|
    units << vrf.terminals
  end
  units.flatten!
  model.getZoneHVACTerminalUnitVariableRefrigerantFlows.each do |t|
    next if units.include? t
    puts "Removing #{t.name}"
    t.remove
  end
end

def low_coordinates(model)

  model.getSpaces.each do |s|
    
  end
end

def spaces_no_surfaces(model)
  model.getSpaces.each do |space|
    next unless space.surfaces.empty?

    puts space.name
  end
end

def change_skylights_to_surfaces(model)
  model.getSubSurfaces.each do |ss|
    next unless ss.subSurfaceType == 'Skylight'
    
    next if ss.space.empty?

    space = ss.space.get
    vertex = ss.vertices
    
    s = OpenStudio::Model::Surface.new(vertex,model)
    s.setSpace(space)
    puts ss.name
    ss.remove

  end
end

def surface_matching(model,spaces)
      # matched surface counter
      initialMatchedSurfaceCounter = 0
      surfaces = model.getSurfaces
      surfaces.each do |surface|
        if surface.outsideBoundaryCondition == 'Surface'
          next if !surface.adjacentSurface.is_initialized # don't count as matched if boundary condition is right but no matched object
          initialMatchedSurfaceCounter += 1
        end
      end
  
      # reporting initial condition of model
      puts "The initial model has #{initialMatchedSurfaceCounter} matched surfaces."
  
      # put all of the spaces in the model into a vector

      # intersect surfaces
      # if intersect_surfaces
        OpenStudio::Model.intersectSurfaces(spaces)
        puts 'Intersecting surfaces, this will create additional geometry.'
      # end
  
      # match surfaces for each space in the vector
      OpenStudio::Model.matchSurfaces(spaces)
      runner.registerInfo('Matching surfaces..')
  
      # matched surface counter
      finalMatchedSurfaceCounter = 0
      surfaces.each do |surface|
        if surface.outsideBoundaryCondition == 'Surface'
          finalMatchedSurfaceCounter += 1
        end
      end
  
      # reporting final condition of model
      puts "The final model has #{finalMatchedSurfaceCounter} matched surfaces."

end

def match_infiltration(target_model,seed_model)

end

def humidity_ratio(db_farenheit, wb_farenheit, elevation_in_ft)

  # Wet bulb in Degree Rankin

  rt = wb_farenheit + 459.67

  # Atmos. Pressure in psia

  pt = 14.696 * (1 - 0.0000068753 * elevation_in_ft) ** 5.2559

  c8 = -10440.4
  c9 = -11.29465
  c10 = -0.027022355
  c11 = 0.00001289036
  c12 = -0.000000002478068
  c13 = 6.5459673

  pws = Math.exp(c8 / rt + c9 + c10 * rt + c11 * rt ** 2 + c12 * rt ** 3 + c13 * Math.log(rt))

  wsat = (pws * 0.62198) / (pt - pws)

  wnom = (1093 - 0.556 * wb_farenheit) * wsat - 0.24 * (db_farenheit - wb_farenheit)
  wdenom = 1093 + 0.444 * db_farenheit - wb_farenheit
  humidity_ratio_result = wnom / wdenom
  return humidity_ratio_result
end

def replace_stp_man_reset(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  model.getAirLoopHVACs.each do |air_loop_hvac|
    node = air_loop_hvac.supplyOutletNode
    node.setpointManagers.each(&:remove) 
    std.air_loop_hvac_enable_supply_air_temperature_reset_warmest_zone(air_loop_hvac)
  end
end

def replace_stp_man_oa_reset(sat_at_lo_oat_c, lo_oat_c, sat_at_hi_oat_c, hi_oat_c, model)
  sat_at_lo_oat_c_si = OpenStudio.convert(sat_at_lo_oat_c, 'F', 'C').get
  lo_oat_c_si = OpenStudio.convert(lo_oat_c, 'F', 'C').get
  sat_at_hi_oat_c_si = OpenStudio.convert(sat_at_hi_oat_c, 'F', 'C').get
  hi_oat_c_si = OpenStudio.convert(hi_oat_c, 'F', 'C').get
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  model.getAirLoopHVACs.each do |air_loop_hvac|
    node = air_loop_hvac.supplyOutletNode
    node.setpointManagers.each(&:remove) 
    # Create a setpoint manager
    sat_oa_reset = OpenStudio::Model::SetpointManagerOutdoorAirReset.new(air_loop_hvac.model)
    sat_oa_reset.setName("#{air_loop_hvac.name} SAT Reset")
    sat_oa_reset.setControlVariable('Temperature')
    sat_oa_reset.setSetpointatOutdoorLowTemperature(sat_at_lo_oat_c_si)
    sat_oa_reset.setOutdoorLowTemperature(lo_oat_c_si)
    sat_oa_reset.setSetpointatOutdoorHighTemperature(sat_at_hi_oat_c_si)
    sat_oa_reset.setOutdoorHighTemperature(hi_oat_c_si)

    # Attach the setpoint manager to the
    # supply outlet node of the system.
    sat_oa_reset.addToNode(air_loop_hvac.supplyOutletNode)
  end
end

# replace_stp_man_oa_reset(61, 30, 51, 50, model)

def economizer_diff_drybulb(max_t_ip, min_t_ip, model)
  max_t_si = OpenStudio.convert(max_t_ip, 'F', 'C').get
  min_t_si = OpenStudio.convert(min_t_ip, 'F', 'C').get
  model.getControllerOutdoorAirs.each do |control|
    control.setEconomizerControlType('FixedDryBulb')
    control.setEconomizerMaximumLimitDryBulbTemperature(max_t_si)
    control.setEconomizerMinimumLimitDryBulbTemperature(min_t_si)
  end
end

def remove_thermostats_from_zones_wo_zonehvac(model)
  model.getThermalZones.each do |z|
    next unless z.equipment.size == 0

    next if z.thermostatSetpointDualSetpoint.empty?

    puts "Deleting from #{z.name}"
    z.thermostatSetpointDualSetpoint.get.remove
  end
end

def remove_humidistats_from_zones_wo_zonehvac(model)
  model.getThermalZones.each do |z|
    next unless z.equipment.size == 0

    next if z.zoneControlHumidistat.empty?

    puts "Deleting from #{z.name}"
    z.zoneControlHumidistat.get.remove
  end
end

# def reduce_DSOA(model,percentage)
#   model.getDesignSpecificationOutdoorAirs.each do |dsoa|
#     puts dsoa.outdoorAirFlowRate
#     fraction_area = dsoa
#     fraction_people = 
#     dsoa.setOutdoorAirFlowperPerson(dsoa.outdoorAirFlowperPerson*percentage)
#     dsoa.setOutdoorAirFlowperFloorArea(dsoa.outdoorAirFlowperFloorArea*percentage)
#     puts ds 
#   end
#   report_oa(model)
# end

def reduce_DSOA(model,percentage)
  # space = model.getSpaceByName('3 - Theater').get
  model.getSpaces.each do |space|
    next if space.designSpecificationOutdoorAir.empty?

    puts space.name
    dsoa = space.designSpecificationOutdoorAir.get
    # area = space.floorArea
    # puts area
    # people = space.numberOfPeople 
    # puts people
    # oa = 0
    # area_oa = dsoa.outdoorAirFlowperFloorArea * area
    # puts area_oa
    # people_oa = dsoa.outdoorAirFlowperPerson  *  people
    # puts people_oa
    # oa = oa + area_oa
    # oa = oa + people_oa
    # next if oa == 0
    # puts oa
    # fraction_area = area_oa / oa
    # puts fraction_area
    # fraction_people = people_oa / oa
    # puts fraction_people
    # percentage = percentage.to_f
    # puts dsoa.outdoorAirFlowperPerson
    # puts dsoa.outdoorAirFlowperPerson * percentage
    # puts dsoa.outdoorAirFlowperFloorArea
    # puts dsoa.outdoorAirFlowperFloorArea *  percentage
    dsoa.setOutdoorAirFlowperPerson(dsoa.outdoorAirFlowperPerson  / percentage)
    dsoa.setOutdoorAirFlowperFloorArea(dsoa.outdoorAirFlowperFloorArea  / percentage)
  end
  report_oa(model)
end

def increase_DSOA(spaces,model,percentage)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
  # space = model.getSpaceByName('3 - Theater').get
  model.getSpaces.each do |space|
    next if space.designSpecificationOutdoorAir.empty?

    # next unless spaces.include? space

    dsoa = space.designSpecificationOutdoorAir.get
    area = space.floorArea
    # puts area
    people = space.numberOfPeople 
    # puts people
    oa = 0
    area_oa = dsoa.outdoorAirFlowperFloorArea * area
    # puts area_oa
    people_oa = dsoa.outdoorAirFlowperPerson  *  people
    # puts people_oa
    oa = oa + area_oa
    oa = oa + people_oa
    next if oa == 0
    # puts oa
    fraction_area = area_oa / oa
    # puts fraction_area
    fraction_people = people_oa / oa
    # puts fraction_people
    percentage = percentage.to_f
    unless people == 0
    end

    dsoa.setOutdoorAirFlowperPerson( percentage *  dsoa.outdoorAirFlowperPerson)

    dsoa.setOutdoorAirFlowperFloorArea(percentage *  dsoa.outdoorAirFlowperFloorArea)
  end
  report_oa(model)
end

def baseline_damper_positions(model,sizing_run_dir = Dir.pwd)
  # attach_sql(model, sql_path)
  # close_sql(model)
  template = '90.1-2010' # or whatever
  std = Standard.build(template)
      # Run sizing run with the HVAC equipment
  std.model_run_sizing_run(model, "#{sizing_run_dir}/SR1")

  std.model_apply_multizone_vav_outdoor_air_sizing(model)
  model.getAirLoopHVACs.sort.each do |air_loop|
    std.air_loop_hvac_apply_minimum_vav_damper_positions(air_loop, false)
  end
  close_sql(model)
end

def print_zone_names(model)
  model.getThermalZones.each do |z|
    puts z.name
  end
end

def rename_lights(model)
  definitions = []
  model.getSpaceTypes.each do |st|
    name = st.name
    puts "#{name} has #{st.lights.size} instances"
    st.lights.each do |light|
      definition = light.lightsDefinition
      definition.setName("Remove")
      definitions << definition
      new_definition = definition.clone(model)
      new_definition = new_definition.to_actual_object
      new_definition.setName("#{name} Lights Definition")
      light.setDefinition(new_definition)
    end
  end
end

def change_surface_height(change_in_height_ip,space_names,model)
  change_in_height_si = OpenStudio.convert(change_in_height_ip, 'in', 'm').get

  ceiling_height = 0
  spaces = []
  surfaces = []
  space_names.each do |name|
    if model.getSpaceByName(name).empty?
      puts "Couldn't find space #{name}"
      next
    end
    space = model.getSpaceByName(name).get
    surfaces << space.surfaces
  end

  surfaces.flatten!

  surfaces.each do |s|

    next if s.surfaceType != "RoofCeiling"

    og_vertex = s.vertices
    new_vertex = []
    og_vertex.each do |v|
      ceiling_height = v.z
      new_v = OpenStudio::Point3d.new(v.x, v.y, v.z + change_in_height_si)
      new_vertex << new_v
    end
    s.setVertices(new_vertex)
    puts s
  end

  surfaces.each do |s|

    og_vertex = s.vertices
    new_vertex = []
    og_vertex.each do |v|
      difference = ceiling_height - v.z
      if difference.abs < 0.1
        new_v = OpenStudio::Point3d.new(v.x, v.y, v.z + change_in_height_si)
        new_vertex << new_v
      else
        new_vertex << v
      end
    end
    s.setVertices(new_vertex)
    puts s
  end
end

def model_remove_unused_resource_objects(model)
  start_size = model.objects.size
  objects = model.purgeUnusedResourceObjects
  objects.each do |obj|
    puts "#{obj.name} is unused; it will be removed."
  end
  end_size = model.objects.size
  puts "The model started with #{start_size} objects and finished with #{end_size} objects after removing unused resource objects."
  return true
end

def change_pthp_to_ptac(model)
  template = '90.1-2004' # or whatever
  std = Standard.build(template)
  zones = []
  model.getThermalZones.each do |z|
    next if z.equipment.empty?

    # z.equipment.each { |e| next if e.to_ZoneHVACPackagedTerminalHeatPump.empty? }

    # z.equipment.each(&:remove)

    zones << z
  end
  system_type = 'PTAC'
  main_heat_fuel = 'NaturalGas'
  zone_heat_fuel = 'NaturalGas'
  cool_fuel = 'Electricity'
  std.model_add_prm_baseline_system(model, system_type, main_heat_fuel, zone_heat_fuel, cool_fuel, zones)
end

def count_spaces_with_space_types(model)

  model.getSpaceTypes.sort.each do |st|
    instances = 0
    st.spaces.each do |s|
      next if s.thermalZone.empty?

      multiplier = s.thermalZone.get.multiplier
      instances = instances + multiplier
    end
    puts "#{st.name} has #{instances} spaces"
  end
end

def account_for_doas(model, strategy, low_stp, high_stp)
  low_stp = OpenStudio.convert(low_stp, 'F', 'C').get
  high_stp = OpenStudio.convert(high_stp, 'F', 'C').get
  model.getThermalZones.each do |z|
    next if z.airLoopHVACs.empty?

    sizing = z.sizingZone
    sizing.setAccountforDedicatedOutdoorAirSystem(true)
    sizing.setDedicatedOutdoorAirSystemControlStrategy(strategy)
    sizing.setDedicatedOutdoorAirLowSetpointTemperatureforDesign(low_stp)
    sizing.setDedicatedOutdoorAirHighSetpointTemperatureforDesign(high_stp)
    puts z.name
  end
end

def do_not_account_for_doas(model)
  model.getThermalZones.each do |z|
    sizing = z.sizingZone
    sizing.setAccountforDedicatedOutdoorAirSystem(false)
  end
end

def make_zone_non_conditioned(zone_vector)
  zone_vector.each do |zone|
    opt_thermostat = zone.thermostatSetpointDualSetpoint
    zone_equipment = zone.equipment
    unless opt_thermostat.empty?
      thermostat = opt_thermostat.get
      puts "Removing #{thermostat.name} in zone #{zone.name}"
      thermostat.remove
    end
    zone_equipment.each do |e|
      puts "Removing #{e.name} in zone #{zone.name}"
      e.remove
    end
  end
end

  
def humidity_ratio(db, wb, elevinft)
  rt = wb + 459.67
  # atmos. pressure in psia
  pt = 14.696 * (1 - 0.0000068753 * elevinft) ** 5.2559
  c8 = -10440.4
  c9 = -11.29465
  c10 = -0.027022355
  c11 = 0.00001289036
  c12 = -0.000000002478068
  c13 = 6.5459673
  pws = Math.exp(c8 / rt + c9 + c10 * rt + c11 * rt ** 2 + c12 * rt ** 3 + c13 * Math.log(rt))
  wsat = (pws * 0.62198) / (pt - pws)
  wnom = (1093 - 0.556 * wb) * wsat - 0.24 * (db - wb)
  wdenom = 1093 + 0.444 * db - wb
  humrat = wnom / wdenom
  return humrat.round(5)
end

def enthalpy(db,hr)
  e = 0.24*db+hr*(1061+0.444*db)
  return e.round(5)
end

def model_classes(model)
  uniq_classes = []
  objects = model.modelObjects(true)
  actual_objects = []
  objects.each do |object|
    actual_object = object.to_actual_object
    actual_objects << actual_object
    class_string = actual_object.class.to_s
    next if uniq_classes.include? class_string

    uniq_classes << class_string
  end
  uniq_classes.each do |cla|
    v =   actual_objects.count { |x| cla.include? x.class.to_s }
    puts "Class #{cla} has #{v} objects"
  end
end

def fix_radiant(model)
  model.getZoneHVACLowTempRadiantVarFlows.each do |radiant|
    tz = radiant.thermalZone.get
    thermostat = tz.thermostatSetpointDualSetpoint.get  
    radiant.setTemperatureControlType('MeanAirTemperature')
    clg_sch = thermostat.coolingSetpointTemperatureSchedule.get
    htg_sch = thermostat.heatingSetpointTemperatureSchedule.get
    cc = radiant.coolingCoil.to_CoilCoolingLowTempRadiantVarFlow.get
    hc = radiant.heatingCoil.to_CoilHeatingLowTempRadiantVarFlow.get
    cc.setCoolingControlTemperatureSchedule(clg_sch)
    hc.setHeatingControlTemperatureSchedule(htg_sch)
  end
end

def set_wwr_st(st, wwr = 0.4, offset = 0.0254, application_type = "Above Floor")
  spaces = st.spaces
  heightOffsetFromFloor = nil
  if application_type == "Above Floor"
    heightOffsetFromFloor = true
  else
    heightOffsetFromFloor = false
  end
  spaces.each do |s|
    s.surfaces.each do |surface|
      next if not surface.outsideBoundaryCondition == "Outdoors"

      next if not surface.surfaceType == "Wall"

      new_window = surface.setWindowToWallRatio(wwr, offset, heightOffsetFromFloor)

      if new_window.empty?
        puts ("Unable to set window-to-wall ratio of " + surface.name.get + " to " + wwr.to_s + ".")
      else
        # not fully accurate - Dan to refactor wiggliness out of C++
        actual = new_window.get.grossArea / surface.grossArea
        puts ("Set window-to-wall ratio of " + surface.name.get + " to " + actual.to_s + ".")
        if not (OpenStudio::DoublesRelativeError(wwr,actual) < 1.0E-3)
          puts ("Tried to set window-to-wall ratio of " + surface.name.get + " to " + wwr.to_s + ", but set to " + actual.to_s + " instead.")
        end
      end
    end
  end
end 

def model_add_heat_recovery_chiller(model,
                                    system_name: 'Chilled Water Loop',
                                    chiller_chilled_water_loop: nil,
                                    chiller_condenser_loop: nil,
                                    chiller_hot_water_loop: nil,
                                    num_chillers: 1)
      
      template = 'NREL ZNE Ready 2017' # or whatever
      std = Standard.build(template)
          # Chillers
      clg_tower_objs = model.getCoolingTowerSingleSpeeds
      model.getChillerElectricEIRs.sort.each { |obj| std.chiller_electric_eir_apply_efficiency_and_curves(obj, clg_tower_objs) }
end