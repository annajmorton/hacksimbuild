# add energy recovery
# add more than 1 boiler 


require 'C:\openstudio-3.2.1\Ruby\openstudio.rb'
require 'fileutils'
require 'openstudio-standards'

### db+ functions
require_relative "dbplushvac"
include DBPLUSHVAC

xlsx_path = File.dirname(__FILE__) + '/output.xlsx'
translator = OpenStudio::OSVersion::VersionTranslator.new
model_path = OpenStudio::Path.new(File.dirname(__FILE__) + '/testo.osm')
model = translator.loadModel(model_path)

model = model.get


# make standard object
# they have 2016, but I need to update the version I have installed
template = '90.1-2019'
std = Standard.build(template)


###
### sort thermal zones in condition type based on proposed
###

office_groups = dbp_get_zones_w_hvac_groups(xlsx_path, 1, model)
#office_tz = dbp_get_zones(xlsx_path, 1, model)
lfs_groups = dbp_get_zones_w_hvac_groups(xlsx_path, 2, model)
#lfs_tz = dbp_get_zones(xlsx_path, 2, model)
lfs_amenity_tz = dbp_get_zones(xlsx_path, 3, model)
of_retail_tz = dbp_get_zones(xlsx_path, 4, model)
lfs_retail_tz = dbp_get_zones(xlsx_path, 5, model)
hwuh_tz = dbp_get_zones(xlsx_path, 6, model)
euh_tz = dbp_get_zones(xlsx_path, 7, model)


###
### remove HVAC systems
###
#remove heating and cooling zone equipment 
std.remove_HVAC(model)
std.remove_all_zone_equipment(model)


###
### create new HVAC systems
###
# add water loops

ls_cw_loop = std.model_add_cw_loop(model,
    system_name: 'LS Condenser Water Loop',
    cooling_tower_type: 'Open Cooling Tower',
    cooling_tower_fan_type: 'Propeller or Axial',
    cooling_tower_capacity_control: 'Variable Speed Fan',
    number_of_cells_per_tower: 2,
    number_cooling_towers: 2,
    use_90_1_design_sizing: true,
    sup_wtr_temp: 70.0,
    dsgn_sup_wtr_temp: 85.0,
    dsgn_sup_wtr_temp_delt: 10.0,
    wet_bulb_approach: 7.0,
    pump_spd_ctrl: 'Variable',
    pump_tot_hd: 49.7)


# Creates a hot water loop with a boiler, district heating, or a
# water-to-water heat pump and adds it to the model.
#
# @param boiler_fuel_type [String] valid choices are Electricity, NaturalGas, PropaneGas, FuelOilNo1, FuelOilNo2, DistrictHeating, HeatPump
# @param ambient_loop [OpenStudio::Model::PlantLoop] The condenser loop for the heat pump. Only used when boiler_fuel_type is HeatPump.
# @param system_name [String] the name of the system, or nil in which case it will be defaulted
# @param dsgn_sup_wtr_temp [Double] design supply water temperature in degrees Fahrenheit, default 180F
# @param dsgn_sup_wtr_temp_delt [Double] design supply-return water temperature difference in degrees Rankine, default 20R
# @param pump_spd_ctrl [String] pump speed control type, Constant or Variable (default)
# @param pump_tot_hd [Double] pump head in ft H2O
# @param boiler_draft_type [String] Boiler type Condensing, MechanicalNoncondensing, Natural (default)
# @param boiler_eff_curve_temp_eval_var [String] LeavingBoiler or EnteringBoiler temperature for the boiler efficiency curve
# @param boiler_lvg_temp_dsgn [Double] boiler leaving design temperature in degrees Fahrenheit
# @param boiler_out_temp_lmt [Double] boiler outlet temperature limit in degrees Fahrenheit
# @param boiler_max_plr [Double] boiler maximum part load ratio
# @param boiler_sizing_factor [Double] boiler oversizing factor
# @return [OpenStudio::Model::PlantLoop] the resulting hot water loop
ls_hw_loop = std.model_add_hw_loop(model,
    'NaturalGas',
    ambient_loop: nil,
    system_name: 'LS Hot Water Loop',
    dsgn_sup_wtr_temp: 180.0,
    dsgn_sup_wtr_temp_delt: 20.0,
    pump_spd_ctrl: 'Variable',
    pump_tot_hd: nil,
    boiler_draft_type: 'Condensing',
    boiler_eff_curve_temp_eval_var: nil,
    boiler_lvg_temp_dsgn: nil,
    boiler_out_temp_lmt: nil,
    boiler_max_plr: nil,
    boiler_sizing_factor: nil)

ls_chw_loop = std.model_add_chw_loop(model,
    system_name: 'LS Chilled Water Loop',
    cooling_fuel: 'Electricity',
    dsgn_sup_wtr_temp: 40.0,
    dsgn_sup_wtr_temp_delt: 10.1,
    chw_pumping_type: 'const_pri_var_sec',
    chiller_cooling_type: "WaterCooled",
    chiller_condenser_type: nil,
    chiller_compressor_type: 'Centrifugal',
    num_chillers: 4,
    condenser_water_loop: ls_cw_loop,
    waterside_economizer: 'non-integrated')


office_cw_loop = std.model_add_cw_loop(model,
    system_name: 'Office Condenser Water Loop',
    cooling_tower_type: 'Open Cooling Tower',
    cooling_tower_fan_type: 'Propeller or Axial',
    cooling_tower_capacity_control: 'Variable Speed Fan',
    number_of_cells_per_tower: 2,
    number_cooling_towers: 2,
    use_90_1_design_sizing: true,
    sup_wtr_temp: 70.0,
    dsgn_sup_wtr_temp: 85.0,
    dsgn_sup_wtr_temp_delt: 10.0,
    wet_bulb_approach: 7.0,
    pump_spd_ctrl: 'Variable',
    pump_tot_hd: 49.7)

office_chw_loop = std.model_add_chw_loop(model,
    system_name: 'Office Chilled Water Loop',
    cooling_fuel: 'Electricity',
    dsgn_sup_wtr_temp: 40.0,
    dsgn_sup_wtr_temp_delt: 10.1,
    chw_pumping_type: 'const_pri_var_sec',
    chiller_cooling_type: "WaterCooled",
    chiller_condenser_type: nil,
    chiller_compressor_type: 'Centrifugal',
    num_chillers: 4,
    condenser_water_loop: office_cw_loop,
    waterside_economizer: 'none')


####
### Systems and Zone HVAC
####

# OFFICE
# central ahus with electric reheat
# std.model_add_pvav_pfp_boxes(model,
#     office_tz,
#     system_name: "Office",
#     chilled_water_loop: office_chw_loop,
#     hvac_op_sch: nil,
#     oa_damper_sch: nil,
#     fan_efficiency: 0.62,
#     fan_motor_efficiency: 0.9,
#     fan_pressure_rise: 4.0)

office_groups.each do |group|
    std.model_add_pvav_pfp_boxes(model,
        group[1],
        system_name: "Office AHU #{group[0]}",
        chilled_water_loop: office_chw_loop,
        hvac_op_sch: nil,
        oa_damper_sch: nil,
        fan_efficiency: 0.62,
        fan_motor_efficiency: 0.9,
        fan_pressure_rise: 4.0)
end


### add retail
std.model_add_pvav_pfp_boxes(model,
    of_retail_tz,
    system_name: "Office Retail",
    chilled_water_loop: office_chw_loop,
    hvac_op_sch: nil,
    oa_damper_sch: nil,
    fan_efficiency: 0.62,
    fan_motor_efficiency: 0.9,
    fan_pressure_rise: 4.0)


# LFS
# fan coils with DOAS on central plant
lfs_groups.each do |group|
    std.model_add_doas_cold_supply(model,
        group[1],
        system_name: "LFS DOAS #{group[0]}",
        hot_water_loop: ls_hw_loop,
        chilled_water_loop: ls_chw_loop,
        hvac_op_sch: nil,
        min_oa_sch: nil,
        min_frac_oa_sch: nil,
        fan_maximum_flow_rate: nil,
        econo_ctrl_mthd: 'FixedDryBulb',
        energy_recovery: TRUE,
        doas_control_strategy: 'NeutralSupplyAir',
        clg_dsgn_sup_air_temp: 45.0,
        htg_dsgn_sup_air_temp: 60.0)
    
    # @param capacity_control_method [String] Capacity control method for the fan coil. Options are ConstantFanVariableFlow,
    #   CyclingFan, VariableFanVariableFlow, and VariableFanConstantFlow.  If VariableFan, the fan will be VariableVolume.
    std.model_add_four_pipe_fan_coil(model,
        group[1],
        ls_chw_loop,
        hot_water_loop: ls_hw_loop,
        ventilation: false)
end



# LFS AMENITY 
# fan power boxes, central AHUs, and hot water reheat
std.model_add_vav_reheat(model,
    lfs_amenity_tz,
    system_name: "LFS Amenity",
    return_plenum: nil,
    heating_type: 'DistrictHeating',
    reheat_type: 'Electric',
    hot_water_loop: ls_hw_loop,
    chilled_water_loop: ls_chw_loop,
    hvac_op_sch: nil,
    oa_damper_sch: nil,
    fan_efficiency: 0.62,
    fan_motor_efficiency: 0.9,
    fan_pressure_rise: 4.0,
    min_sys_airflow_ratio: 0.3,
    vav_sizing_option: 'Coincident',
    econo_ctrl_mthd: nil)

# LFS Retail
std.model_add_vav_reheat(model,
    lfs_retail_tz,
    system_name: "LFS Retail",
    return_plenum: nil,
    heating_type: 'DistrictHeating',
    reheat_type: 'Electric',
    hot_water_loop: ls_hw_loop,
    chilled_water_loop: ls_chw_loop,
    hvac_op_sch: nil,
    oa_damper_sch: nil,
    fan_efficiency: 0.62,
    fan_motor_efficiency: 0.9,
    fan_pressure_rise: 4.0,
    min_sys_airflow_ratio: 0.3,
    vav_sizing_option: 'Coincident',
    econo_ctrl_mthd: nil)

# HWUH
fanpressure = 0.2
std.model_add_unitheater(model,
    hwuh_tz,
    hvac_op_sch: nil,
    fan_control_type: 'OnOff',
    fan_pressure_rise: fanpressure,
    heating_type: "DistrictHeating",
    hot_water_loop: ls_hw_loop,
    rated_inlet_water_temperature: 180.0,
    rated_outlet_water_temperature: 160.0,
    rated_inlet_air_temperature: 60.0,
    rated_outlet_air_temperature: 104.0)

# EUH
fanpressure = 0.2
std.model_add_unitheater(model,euh_tz,fan_control_type:"OnOff",fan_pressure_rise:fanpressure,heating_type:"Electric")



# add elevators
# elev_spaces = dbp_get_spaces_from_zones(elev_tz)
# elevator_schedule = std.spaces_get_occupancy_schedule(elev_spaces)
# puts elevator_schedule

# std.model_add_schedule(model,elevator_schedule.name.get.to_s)

# elev_spaces.each do |elev_space|
#     std.model_add_elevator(model,
#         elev_space,
#         6,
#         "Traction",
#         elevator_schedule,
#         elevator_schedule,
#         elevator_schedule,
#         building_type = nil)
# end 

#std.model_add_elevators(model)




###
### Save model
###

model.save(model_path,1)