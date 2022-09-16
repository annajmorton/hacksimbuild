require "rubyXL"
require 'rubyXL/convenience_methods/cell'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL/convenience_methods/worksheet'


### modified versions
### of the openstudio havc functions 
module DBPLUSHVAC

    #### 
    ### LOAD the Zones
    ###
    ###
    def dbp_get_zones(xlsx_path, sheet_no, model)
        workbook = RubyXL::Parser.parse(xlsx_path)
        tz = []
            workbook.worksheets[sheet_no].each { |row|
            if !row.nil?
                tzname = row[0].value
        
                if tzname ==   "Thermal Zone"
                    next
                end
        
                if !tzname.nil? 
                    if !tzname.empty?
                        zone = model.getThermalZoneByName(tzname).get
                        tz.push(zone)
                    end
                end
            end 
        }
        return tz
    end
    
    ### get the zones
    ### second column is the HVAC group
    ### returns a hash
    def dbp_get_zones_w_hvac_groups(xlsx_path, sheet_no, model)
        workbook = RubyXL::Parser.parse(xlsx_path)
        hvac_group = {}
        workbook.worksheets[sheet_no].each { |row|
            if !row.nil?
                tzname = row[0].value
                group = row[1].value
    
                if tzname == "Thermal Zone"
                    next
                end
    
                if !tzname.nil? 
                    if !tzname.empty?
                        zone = model.getThermalZoneByName(tzname).get
                        puts tzname
                        if !hvac_group[group].nil?
                            hvac_group[group] = [zone] + hvac_group[group] 
                        else 
                            hvac_group[group] = [zone]
                        end
                    
                    end
                end
            end 
        }
        return hvac_group 
    end

    ## get spaces 
    ## from list of Thermal Zones
    def dbp_get_spaces_from_zones(thermal_zones)
        spaces = []
        thermal_zones.each do |tz|
            tz.spaces.each do |space|
                spaces.push(space)
            end
        end

        return spaces
    end
    
    # this function adds a desuper heater
    # it doesn't impact compressor efficiency 
    # this is hot-gas reheat
    # heat_reclaim_recovery_efficiency is in ratio
    def model_add_coilheatingdesuperheater(std,
        model,
        heat_reclaim_recovery_efficiency,
        air_loop)
        
        air_loop.supplyComponents.each do |supply_component|
            if not supply_component.to_CoilCoolingDXSingleSpeed.empty?
                #get the cooling coil
                dx_coil = supply_component.to_CoilCoolingDXSingleSpeed.get 
                #create a CoilHeatingDesuperheater  and add it infront of the Cooling Coil
                chd = OpenStudio::Model::CoilHeatingDesuperheater.new(model)  
                dxcc_node = dx_coil.outletModelObject.get.to_Node
                if not dxcc_node.empty?
                    #set node connection
                    chd.addToNode(dxcc_node.get)
                    #set required fields to a single default value
                    chd.setHeatReclaimRecoveryEfficiency(heat_reclaim_recovery_efficiency)
                    chd.setHeatingSource(dx_coil)

                    avail_sch = model.alwaysOnDiscreteSchedule
                    chd.setAvailabilitySchedule(avail_sch)
                    puts "desuperheater added"
                end
            end	
        end
    end
###
  # Creates a DOAS system with terminal units for each zone.
  #
  # @param thermal_zones [Array<OpenStudio::Model::ThermalZone>] array of zones to connect to this system
  # @param system_name [String] the name of the system, or nil in which case it will be defaulted
  # @param doas_type [String] DOASCV or DOASVAV, determines whether the DOAS is operated at scheduled,
  #   constant flow rate, or airflow is variable to allow for economizing or demand controlled ventilation
  # @param doas_control_strategy [String] DOAS control strategy
  # @param hot_water_loop [OpenStudio::Model::PlantLoop] hot water loop to connect to heating and zone fan coils
  # @param chilled_water_loop [OpenStudio::Model::PlantLoop] chilled water loop to connect to cooling coil
  # @param condenser_water_loop [OpenStudio::Model::PlantLoop] chilled water loop to connect to cooling coil
  # @param hvac_op_sch [String] name of the HVAC operation schedule, default is always on
  # @param min_oa_sch [String] name of the minimum outdoor air schedule, default is always on
  # @param min_frac_oa_sch [String] name of the minimum fraction of outdoor air schedule, default is always on
  # @param fan_maximum_flow_rate [Double] fan maximum flow rate in cfm, default is autosize
  # @param econo_ctrl_mthd [String] economizer control type, default is Fixed Dry Bulb
  #   If enabled, the DOAS will be sized for twice the ventilation minimum to allow economizing
  # @param include_exhaust_fan [Bool] if true, include an exhaust fan
  # @param clg_dsgn_sup_air_temp [Double] design cooling supply air temperature in degrees Fahrenheit, default 65F
  # @param htg_dsgn_sup_air_temp [Double] design heating supply air temperature in degrees Fahrenheit, default 75F
  # @return [OpenStudio::Model::AirLoopHVAC] the resulting DOAS air loop

    def dbp_model_add_doas(std,
            model,
            thermal_zones,
            system_name: nil,
            doas_type: 'DOASCV',
            hot_water_loop: nil,
            chilled_water_loop: nil,
            condenser_water_loop: nil,
            hvac_op_sch: nil,
            min_oa_sch: nil,
            min_frac_oa_sch: nil,
            fan_maximum_flow_rate: nil,
            econo_ctrl_mthd: 'NoEconomizer',
            include_exhaust_fan: true,
            demand_control_ventilation: false,
            doas_control_strategy: 'NeutralSupplyAir',
            clg_dsgn_sup_air_temp: 60.0,
            htg_dsgn_sup_air_temp: 70.0)

        # Check the total OA requirement for all zones on the system
        tot_oa_req = 0
        thermal_zones.each do |zone|
            tot_oa_req += std.thermal_zone_outdoor_airflow_rate(zone)
        end

        # If the total OA requirement is zero do not add the DOAS system because the simulations will fail
        if tot_oa_req.zero?
        OpenStudio.logFree(OpenStudio::Info, 'openstudio.Model.Model', "Not adding DOAS system for #{thermal_zones.size} zones because combined OA requirement for all zones is zero.")
        return false
        end
        OpenStudio.logFree(OpenStudio::Info, 'openstudio.Model.Model', "Adding DOAS system for #{thermal_zones.size} zones.")

        # create a DOAS air loop
        air_loop = OpenStudio::Model::AirLoopHVAC.new(model)
        if system_name.nil?
        air_loop.setName("#{thermal_zones.size} Zone DOAS")
        else
        air_loop.setName(system_name)
        end

        # set availability schedule
        if hvac_op_sch.nil?
        hvac_op_sch = model.alwaysOnDiscreteSchedule
        else
        hvac_op_sch = model_add_schedule(model, hvac_op_sch)
        end

        # DOAS design temperatures
        if clg_dsgn_sup_air_temp.nil?
        clg_dsgn_sup_air_temp_c = OpenStudio.convert(60.0, 'F', 'C').get
        else
        clg_dsgn_sup_air_temp_c = OpenStudio.convert(clg_dsgn_sup_air_temp, 'F', 'C').get
        end

        if htg_dsgn_sup_air_temp.nil?
        htg_dsgn_sup_air_temp_c = OpenStudio.convert(70.0, 'F', 'C').get
        else
        htg_dsgn_sup_air_temp_c = OpenStudio.convert(htg_dsgn_sup_air_temp, 'F', 'C').get
        end

        # modify system sizing properties
        sizing_system = air_loop.sizingSystem
        sizing_system.setTypeofLoadtoSizeOn('VentilationRequirement')
        sizing_system.setAllOutdoorAirinCooling(true)
        sizing_system.setAllOutdoorAirinHeating(true)
        # set minimum airflow ratio to 1.0 to avoid under-sizing heating coil
        if model.version < OpenStudio::VersionString.new('2.7.0')
        sizing_system.setMinimumSystemAirFlowRatio(1.0)
        else
        sizing_system.setCentralHeatingMaximumSystemAirFlowRatio(1.0)
        end
        sizing_system.setSizingOption('Coincident')
        sizing_system.setCentralCoolingDesignSupplyAirTemperature(clg_dsgn_sup_air_temp_c)
        sizing_system.setCentralHeatingDesignSupplyAirTemperature(htg_dsgn_sup_air_temp_c)

        if doas_type == 'DOASCV'
        supply_fan = std.create_fan_by_name(model,
                                'Constant_DOAS_Fan',
                                fan_name: 'DOAS Supply Fan',
                                end_use_subcategory: 'DOAS Fans')
        else # 'DOASVAV'
        supply_fan = create_fan_by_name(model,
                                'Variable_DOAS_Fan',
                                fan_name: 'DOAS Supply Fan',
                                end_use_subcategory: 'DOAS Fans')
        end
        supply_fan.setAvailabilitySchedule(model.alwaysOnDiscreteSchedule)
        supply_fan.setMaximumFlowRate(OpenStudio.convert(fan_maximum_flow_rate, 'cfm', 'm^3/s').get) unless fan_maximum_flow_rate.nil?
        supply_fan.addToNode(air_loop.supplyInletNode)

        # create heating coil
        if !condenser_water_loop.nil? 
            std.create_coil_heating_water_to_air_heat_pump_equation_fit(model,
                condenser_water_loop,
                air_loop_node: air_loop.supplyInletNode,
                name: "#{air_loop.name} Water-to-Air HP Htg Coil",
                type: nil,
                cop: 3.5)

        elsif hot_water_loop.nil?
            # electric backup heating coil
            create_coil_heating_electric(model,
                                air_loop_node: air_loop.supplyInletNode,
                                name: "#{air_loop.name} Backup Htg Coil")
            # heat pump coil
            create_coil_heating_dx_single_speed(model,
                                        air_loop_node: air_loop.supplyInletNode,
                                        name: "#{air_loop.name} Htg Coil")
        else
            create_coil_heating_water(model,
                            hot_water_loop,
                            air_loop_node: air_loop.supplyInletNode,
                            name: "#{air_loop.name} Htg Coil",
                            controller_convergence_tolerance: 0.0001)
        end

        # could add a humidity controller here set to limit supply air to a 16.6C/62F dewpoint
        # the default outdoor air reset to 60F prevents exceeding this dewpoint in all ASHRAE climate zones
        # the humidity controller needs a DX coil that can control humidity, e.g. CoilCoolingDXTwoStageWithHumidityControlMode
        # max_humidity_ratio_sch = model_add_constant_schedule_ruleset(model,
        #                                                              0.012,
        #                                                              name = "0.012 Humidity Ratio Schedule",
        #                                                              sch_type_limit: "Humidity Ratio")
        # sat_oa_reset = OpenStudio::Model::SetpointManagerScheduled.new(model, max_humidity_ratio_sch)
        # sat_oa_reset.setName("#{air_loop.name.to_s} Humidity Controller")
        # sat_oa_reset.setControlVariable('MaximumHumidityRatio')
        # sat_oa_reset.addToNode(air_loop.supplyInletNode)

        # create cooling coil
        if !condenser_water_loop.nil?
            std.create_coil_cooling_water_to_air_heat_pump_equation_fit(model,
                condenser_water_loop,
                air_loop_node: air_loop.supplyInletNode,
                name: "#{air_loop.name} Water-to-Air HP Clg Coil",
                type: nil,
                cop: 3.5)

        elsif chilled_water_loop.nil?
        create_coil_cooling_dx_two_speed(model,
                                air_loop_node: air_loop.supplyInletNode,
                                name: "#{air_loop.name} 2spd DX Clg Coil",
                                type: 'OS default')
        else
        create_coil_cooling_water(model,
                        chilled_water_loop,
                        air_loop_node: air_loop.supplyInletNode,
                        name: "#{air_loop.name} Clg Coil")
        end

        # minimum outdoor air schedule
        unless min_oa_sch.nil?
        min_oa_sch = model_add_schedule(model, min_oa_sch)
        end

        # minimum outdoor air fraction schedule
        if min_frac_oa_sch.nil?
        min_frac_oa_sch = model.alwaysOnDiscreteSchedule
        else
        min_frac_oa_sch = model_add_schedule(model, min_frac_oa_sch)
        end

        # create controller outdoor air
        controller_oa = OpenStudio::Model::ControllerOutdoorAir.new(model)
        controller_oa.setName("#{air_loop.name} Outdoor Air Controller")
        controller_oa.setEconomizerControlType(econo_ctrl_mthd)
        controller_oa.setMinimumLimitType('FixedMinimum')
        controller_oa.autosizeMinimumOutdoorAirFlowRate
        controller_oa.setMinimumOutdoorAirSchedule(min_oa_sch) unless min_oa_sch.nil?
        controller_oa.setMinimumFractionofOutdoorAirSchedule(min_frac_oa_sch)
        controller_oa.resetEconomizerMinimumLimitDryBulbTemperature
        controller_oa.resetEconomizerMaximumLimitDryBulbTemperature
        controller_oa.resetEconomizerMaximumLimitEnthalpy
        controller_oa.resetMaximumFractionofOutdoorAirSchedule
        controller_oa.setHeatRecoveryBypassControlType('BypassWhenWithinEconomizerLimits')
        controller_mech_vent = controller_oa.controllerMechanicalVentilation
        controller_mech_vent.setName("#{air_loop.name} Mechanical Ventilation Controller")
        controller_mech_vent.setDemandControlledVentilation(true) if demand_control_ventilation
        controller_mech_vent.setSystemOutdoorAirMethod('ZoneSum')

        # create outdoor air system
        oa_system = OpenStudio::Model::AirLoopHVACOutdoorAirSystem.new(model, controller_oa)
        oa_system.setName("#{air_loop.name} OA System")
        oa_system.addToNode(air_loop.supplyInletNode)

        # create an exhaust fan
        if include_exhaust_fan
        if doas_type == 'DOASCV'
        exhaust_fan = std.create_fan_by_name(model,
                                    'Constant_DOAS_Fan',
                                    fan_name: 'DOAS Exhaust Fan',
                                    end_use_subcategory: 'DOAS Fans')
        else # 'DOASVAV'
        exhaust_fan = create_fan_by_name(model,
                                    'Variable_DOAS_Fan',
                                    fan_name: 'DOAS Exhaust Fan',
                                    end_use_subcategory: 'DOAS Fans')
        end
        # set pressure rise 1.0 inH2O lower than supply fan, 1.0 inH2O minimum
        exhaust_fan_pressure_rise = supply_fan.pressureRise - OpenStudio.convert(1.0, 'inH_{2}O', 'Pa').get
        exhaust_fan_pressure_rise = OpenStudio.convert(1.0, 'inH_{2}O', 'Pa').get if exhaust_fan_pressure_rise < OpenStudio.convert(1.0, 'inH_{2}O', 'Pa').get
        exhaust_fan.setPressureRise(exhaust_fan_pressure_rise)
        exhaust_fan.addToNode(air_loop.supplyInletNode)
        end

        # create a setpoint manager
        sat_oa_reset = OpenStudio::Model::SetpointManagerOutdoorAirReset.new(model)
        sat_oa_reset.setName("#{air_loop.name} SAT Reset")
        sat_oa_reset.setControlVariable('Temperature')
        sat_oa_reset.setSetpointatOutdoorLowTemperature(htg_dsgn_sup_air_temp_c)
        sat_oa_reset.setOutdoorLowTemperature(OpenStudio.convert(55.0, 'F', 'C').get)
        sat_oa_reset.setSetpointatOutdoorHighTemperature(clg_dsgn_sup_air_temp_c)
        sat_oa_reset.setOutdoorHighTemperature(OpenStudio.convert(70.0, 'F', 'C').get)
        sat_oa_reset.addToNode(air_loop.supplyOutletNode)

        # set air loop availability controls and night cycle manager, after oa system added
        air_loop.setAvailabilitySchedule(hvac_op_sch)
        air_loop.setNightCycleControlType('CycleOnAnyZoneFansOnly')

        # add thermal zones to airloop
        thermal_zones.each do |zone|
        # skip zones with no outdoor air flow rate
        unless std.thermal_zone_outdoor_airflow_rate(zone) > 0
        OpenStudio.logFree(OpenStudio::Debug, 'openstudio.Model.Model', "---#{zone.name} has no outdoor air flow rate and will not be added to #{air_loop.name}")
        next
        end

        OpenStudio.logFree(OpenStudio::Debug, 'openstudio.Model.Model', "---adding #{zone.name} to #{air_loop.name}")

        # make an air terminal for the zone
        if doas_type == 'DOASCV'
        air_terminal = OpenStudio::Model::AirTerminalSingleDuctUncontrolled.new(model, model.alwaysOnDiscreteSchedule)
        elsif doas_type == 'DOASVAVReheat'
        # Reheat coil
        if hot_water_loop.nil?
        rht_coil = create_coil_heating_electric(model, name: "#{zone.name} Electric Reheat Coil")
        else
        rht_coil = create_coil_heating_water(model, hot_water_loop, name: "#{zone.name} Reheat Coil")
        end
        # VAV reheat terminal
        air_terminal = OpenStudio::Model::AirTerminalSingleDuctVAVReheat.new(model, model.alwaysOnDiscreteSchedule, rht_coil)
        if model.version < OpenStudio::VersionString.new('3.0.1')
        air_terminal.setZoneMinimumAirFlowMethod('Constant')
        else
        air_terminal.setZoneMinimumAirFlowInputMethod('Constant')
        end
        air_terminal.setControlForOutdoorAir(true) if demand_control_ventilation
        else # 'DOASVAV'
        air_terminal = OpenStudio::Model::AirTerminalSingleDuctVAVNoReheat.new(model, model.alwaysOnDiscreteSchedule)
        if model.version < OpenStudio::VersionString.new('3.0.1')
        air_terminal.setZoneMinimumAirFlowMethod('Constant')
        else
        air_terminal.setZoneMinimumAirFlowInputMethod('Constant')
        end
        air_terminal.setConstantMinimumAirFlowFraction(0.1)
        air_terminal.setControlForOutdoorAir(true) if demand_control_ventilation
        end
        air_terminal.setName("#{zone.name} Air Terminal")

        # attach new terminal to the zone and to the airloop
        air_loop.multiAddBranchForZone(zone, air_terminal.to_HVACComponent.get)

        # ensure the DOAS takes priority, so ventilation load is included when treated by other zonal systems
        # From EnergyPlus I/O reference:
        # "For situations where one or more equipment types has limited capacity or limited control capability, order the
        #  sequence so that the most controllable piece of equipment runs last. For example, with a dedicated outdoor air
        #  system (DOAS), the air terminal for the DOAS should be assigned Heating Sequence = 1 and Cooling Sequence = 1.
        #  Any other equipment should be assigned sequence 2 or higher so that it will see the net load after the DOAS air
        #  is added to the zone."
        zone.setCoolingPriority(air_terminal.to_ModelObject.get, 1)
        zone.setHeatingPriority(air_terminal.to_ModelObject.get, 1)

        # set the cooling and heating fraction to zero so that if DCV is enabled,
        # the system will lower the ventilation rate rather than trying to meet the heating or cooling load.
        if model.version < OpenStudio::VersionString.new('2.8.0')
        if demand_control_ventilation
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.Model.Model', 'Unable to add DOAS with DCV to model because the setSequentialCoolingFraction method is not available in OpenStudio versions less than 2.8.0.')
        else
        OpenStudio.logFree(OpenStudio::Warn, 'openstudio.Model.Model', 'OpenStudio version is less than 2.8.0.  The DOAS system will not be able to have DCV if changed at a later date.')
        end
        else
        zone.setSequentialCoolingFraction(air_terminal.to_ModelObject.get, 0.0)
        zone.setSequentialHeatingFraction(air_terminal.to_ModelObject.get, 0.0)

        # if economizing, override to meet cooling load first with doas supply
        unless econo_ctrl_mthd == 'NoEconomizer'
        zone.setSequentialCoolingFraction(air_terminal.to_ModelObject.get, 1.0)
        end
        end

        # DOAS sizing
        sizing_zone = zone.sizingZone
        sizing_zone.setAccountforDedicatedOutdoorAirSystem(true)
        sizing_zone.setDedicatedOutdoorAirSystemControlStrategy(doas_control_strategy)
        sizing_zone.setDedicatedOutdoorAirLowSetpointTemperatureforDesign(clg_dsgn_sup_air_temp_c)
        sizing_zone.setDedicatedOutdoorAirHighSetpointTemperatureforDesign(htg_dsgn_sup_air_temp_c)
        end

        return air_loop
    end

    def airloop_water_to_air(std,
                             model,plant_loop,
                             air_loop_node,
                             heatingCOP,
                             coolingCOP)
        std.create_coil_heating_water_to_air_heat_pump_equation_fit(model,
            plant_loop,
            air_loop_node: airloop_loop_node,
            name: 'Water-to-Air HP Htg Coil',
            type: nil,
            cop: heatingCOP)
        std.create_coil_cooling_water_to_air_heat_pump_equation_fit(model,
            plant_loop,
            air_loop_node: nil,
            name: 'Water-to-Air HP Clg Coil',
            type: nil,
            cop: coolingCOP)
    end 

    # this creates a dahu - dehumidification air handling unit
    # these units are used in spaces with very high latent loads
    # this model includes energy recovery 
    # between the return at and the cooling coil 
    # leavig air 
    # typically these units do not have outside air
    #@param min oa and min oa fraction schedules, and economizer are not uses
    def model_add_dahu(std,
        model,
        thermal_zones,
        system_name: nil,
        hot_water_loop: nil,
        chilled_water_loop: nil,
        hvac_op_sch: nil,
        min_oa_sch: nil,
        min_frac_oa_sch: nil,
        fan_maximum_flow_rate: nil,
        econo_ctrl_mthd: 'FixedDryBulb',
        energy_recovery: false,
        dahu_control_strategy: 'NeutralSupplyAir',
        clg_dsgn_sup_air_temp: 55.0,
        htg_dsgn_sup_air_temp: 60.0)



        # create a DAHU air loop
        air_loop = OpenStudio::Model::AirLoopHVAC.new(model)
        if system_name.nil?
            air_loop.setName("#{thermal_zones.size} Zone DAHU")
        else
            air_loop.setName(system_name)
        end

        # set availability schedule
        if hvac_op_sch.nil?
            hvac_op_sch = model.alwaysOnDiscreteSchedule
        else
            hvac_op_sch = model_add_schedule(model, hvac_op_sch)
        end

        # design temperatures
        if clg_dsgn_sup_air_temp.nil?
            clg_dsgn_sup_air_temp_c = OpenStudio.convert(55.0, 'F', 'C').get
        else
            clg_dsgn_sup_air_temp_c = OpenStudio.convert(clg_dsgn_sup_air_temp, 'F', 'C').get
        end

        if htg_dsgn_sup_air_temp.nil?
            htg_dsgn_sup_air_temp_c = OpenStudio.convert(60.0, 'F', 'C').get
        else
            htg_dsgn_sup_air_temp_c = OpenStudio.convert(htg_dsgn_sup_air_temp, 'F', 'C').get
        end

        # modify system sizing properties
        sizing_system = air_loop.sizingSystem
        sizing_system.setTypeofLoadtoSizeOn('Total')
        sizing_system.setAllOutdoorAirinCooling(true)
        sizing_system.setAllOutdoorAirinHeating(true)
        # set minimum airflow ratio to 1.0 to avoid under-sizing heating coil
        if model.version < OpenStudio::VersionString.new('2.7.0')
            sizing_system.setMinimumSystemAirFlowRatio(1.0)
        else
            sizing_system.setCentralHeatingMaximumSystemAirFlowRatio(1.0)
        end
        sizing_system.setSizingOption('Coincident')
        sizing_system.setCentralCoolingDesignSupplyAirTemperature(clg_dsgn_sup_air_temp_c)
        sizing_system.setCentralHeatingDesignSupplyAirTemperature(htg_dsgn_sup_air_temp_c)

        # create supply fan
        supply_fan = std.create_fan_by_name(model,
            'PSZ_VAV_Fan',
            fan_name: 'DAHU Supply Fan',
            end_use_subcategory: 'DAHU Fans')
        supply_fan.setAvailabilitySchedule(model.alwaysOnDiscreteSchedule)
        supply_fan.setMaximumFlowRate(OpenStudio.convert(fan_maximum_flow_rate, 'cfm', 'm^3/s').get) unless fan_maximum_flow_rate.nil?
        supply_fan.addToNode(air_loop.supplyInletNode)

        # create heating coil
        if hot_water_loop.nil?
            # electric backup heating coil
            std.create_coil_heating_electric(model, 
                air_loop_node: air_loop.supplyInletNode,
                name: "#{air_loop.name} Backup Htg Coil")
            # heat pump coil
            std.create_coil_heating_dx_single_speed(model,
                        air_loop_node: air_loop.supplyInletNode,
                        name: "#{air_loop.name} Htg Coil")
        else
            std.create_coil_heating_water(model,
                hot_water_loop,
                air_loop_node: air_loop.supplyInletNode,
                name: "#{air_loop.name} Htg Coil",
                controller_convergence_tolerance: 0.0001)
        end

        # add energy recovery to cooling coil if requested
        if energy_recovery
            if chilled_water_loop.nil?
                #need to add definition for dx
                puts "script not defined for dx system"
            else  
                #Create cooling coil with hx

                clg_ass_coil = OpenStudio::Model::CoilSystemCoolingWaterHeatExchangerAssisted.new(model)
                clg_ass_coil.setName("#{air_loop.name} Coil System Clg HX Assisted")
                clg_ass_coil.addToNode(air_loop.supplyInletNode)

                clg_coil_name = "#{air_loop.name} Clg Coil"
                clg_coil = clg_ass_coil.coolingCoil
    
                # add to chilled water loop
                chilled_water_loop.addDemandBranchForComponent(clg_coil)

                # set coil name
                clg_coil.setName(clg_coil_name)
                

                # create a humidity setpoint manager
                sat_reht_reset = OpenStudio::Model::SetpointManagerSingleZoneReheat.new(model)
                sat_reht_reset.setName("#{air_loop.name} SAT Reheat Reset")
                sat_reht_reset.setControlVariable('Temperature')
                sat_reht_reset.setMinimumSupplyAirTemperature(clg_dsgn_sup_air_temp_c)
                sat_reht_reset.setMaximumSupplyAirTemperature(htg_dsgn_sup_air_temp_c)
                sat_reht_reset.addToNode(air_loop.supplyOutletNode)

                # set coil availability schedule
                #coil_availability_schedule = model.alwaysOnDiscreteSchedule
                #clg_coil.setAvailabilitySchedule(coil_availability_schedule)

                # rated temperatures
                #clg_coil.autosizeDesignInletWaterTemperature

                #clg_coil.setDesignInletAirTemperature(design_inlet_air_temperature) unless design_inlet_air_temperature.nil?
                #clg_coil.setDesignOutletAirTemperature(design_outlet_air_temperature) unless design_outlet_air_temperature.nil?

                # defaults
                #clg_coil.setHeatExchangerConfiguration('CrossFlow')

                # coil controller properties
                # NOTE: These inputs will get overwritten if addToNode or addDemandBranchForComponent is called on the htg_coil object after this
                #clg_coil_controller = clg_coil.controllerWaterCoil.get
                #clg_coil_controller.setName("#{clg_coil.name} Controller")
                #clg_coil_controller.setAction('Reverse')
                #clg_coil_controller.setMinimumActuatedFlow(0.0)


                #erv.addToNode(oa_system.outboardOANode.get)
                #erv.setHeatExchangerType('Rotary')
                # TODO: come up with scheme for estimating power of ERV motor wheel which might require knowing airflow.
                # erv.setNominalElectricPower(value_new)
                #erv.setEconomizerLockout(true)
                #erv.setSupplyAirOutletTemperatureControl(false)

                #erv.setSensibleEffectivenessat100HeatingAirFlow(0.76)
                #erv.setSensibleEffectivenessat75HeatingAirFlow(0.81)
                #erv.setLatentEffectivenessat100HeatingAirFlow(0.68)
                #erv.setLatentEffectivenessat75HeatingAirFlow(0.73)
            
                #erv.setSensibleEffectivenessat100CoolingAirFlow(0.76)
                #erv.setSensibleEffectivenessat75CoolingAirFlow(0.81)
                #erv.setLatentEffectivenessat100CoolingAirFlow(0.68)
                #erv.setLatentEffectivenessat75CoolingAirFlow(0.73)

                # increase fan static pressure to account for ERV
                #erv_pressure_rise = OpenStudio.convert(1.0, 'inH_{2}O', 'Pa').get
                #new_pressure_rise = supply_fan.pressureRise + erv_pressure_rise
                #supply_fan.setPressureRise(new_pressure_rise)
            end 
        else
            # create cooling coil
            if chilled_water_loop.nil?
                clg_coil_name = "#{air_loop.name} 2spd DX Clg Coil"
                std.create_coil_cooling_dx_two_speed(model,
                        air_loop_node: air_loop.supplyInletNode,
                        name: clg_coil_name,
                        type: 'OS default')
            else
                clg_coil_name = "#{air_loop.name} Clg Coil"
                std.create_coil_cooling_water(model,
                    chilled_water_loop,
                    air_loop_node: air_loop.supplyInletNode,
                    name: clg_coil_name)
            end
        end 


        # create a setpoint manager
        sat_reht_reset = OpenStudio::Model::SetpointManagerSingleZoneReheat.new(model)
        sat_reht_reset.setName("#{air_loop.name} SAT Reheat Reset")
        sat_reht_reset.setControlVariable('Temperature')
        sat_reht_reset.setMinimumSupplyAirTemperature(clg_dsgn_sup_air_temp_c)
        sat_reht_reset.setMaximumSupplyAirTemperature(htg_dsgn_sup_air_temp_c)
        sat_reht_reset.addToNode(air_loop.supplyOutletNode)

        # set air loop availability controls and night cycle manager, after oa system added
        air_loop.setAvailabilitySchedule(hvac_op_sch)
        air_loop.setNightCycleControlType('CycleOnAny')

    
        # add thermal zones to airloop
        thermal_zones.each do |zone|
            OpenStudio.logFree(OpenStudio::Debug, 'openstudio.Model.Model', "---adding #{zone.name} to #{air_loop.name}")

            # make an air terminal for the zone
            air_terminal = OpenStudio::Model::AirTerminalSingleDuctUncontrolled.new(model, model.alwaysOnDiscreteSchedule)
            air_terminal.setName("#{zone.name} Air Terminal")

            # attach new terminal to the zone and to the airloop
            air_loop.multiAddBranchForZone(zone, air_terminal.to_HVACComponent.get)

            # DOAS sizing
            sizing_zone = zone.sizingZone
            #sizing_zone.setAccountforDedicatedOutdoorAirSystem(true)
            #sizing_zone.setDedicatedOutdoorAirSystemControlStrategy('ColdSupplyAir')
            #sizing_zone.setDedicatedOutdoorAirLowSetpointTemperatureforDesign(clg_dsgn_sup_air_temp_c)
            #sizing_zone.setDedicatedOutdoorAirHighSetpointTemperatureforDesign(htg_dsgn_sup_air_temp_c)
        end

        return air_loop
    end


  # Creates a PSZ-AC system for each zone and adds it to the model.
  #
  # @param thermal_zones [Array<OpenStudio::Model::ThermalZone>] array of zones to connect to this system
  # @param system_name [String] the name of the system, or nil in which case it will be defaulted
  # @param cooling_type [String] valid choices are Water, Two Speed DX AC, Single Speed DX AC, Single Speed Heat Pump, Water To Air Heat Pump
  # @param chilled_water_loop [OpenStudio::Model::PlantLoop] chilled water loop to connect cooling coil to, or nil
  # @param hot_water_loop [OpenStudio::Model::PlantLoop] hot water loop to connect heating coil to, or nil
  # @param heating_type [String] valid choices are NaturalGas, Electricity, Water, Single Speed Heat Pump, Water To Air Heat Pump, or nil (no heat)
  # @param supplemental_heating_type [String] valid choices are Electricity, NaturalGas,  nil (no heat)
  # @param fan_location [String] valid choices are BlowThrough, DrawThrough
  # @param fan_type [String] valid choices are ConstantVolume, Cycling
  # @param hvac_op_sch [String] name of the HVAC operation schedule or nil in which case will be defaulted to always on
  # @param oa_damper_sch [String] name of the oa damper schedule or nil in which case will be defaulted to always open
  # @return [Array<OpenStudio::Model::AirLoopHVAC>] an array of the resulting PSZ-AC air loops


    def model_add_gpsz_ac(std,
        model,
        thermal_zones,
        system_name: nil,
        cooling_type: 'Single Speed DX AC',
        chilled_water_loop: nil,
        hot_water_loop: nil,
        heating_type: nil,
        supplemental_heating_type: nil,
        fan_location: 'DrawThrough',
        fan_type: 'ConstantVolume',
        hvac_op_sch: nil,
        oa_damper_sch: nil)

        # hvac operation schedule
        if hvac_op_sch.nil?
        hvac_op_sch = model.alwaysOnDiscreteSchedule
        else
        hvac_op_sch = model_add_schedule(model, hvac_op_sch)
        end

        # oa damper schedule
        if oa_damper_sch.nil?
        oa_damper_sch = model.alwaysOnDiscreteSchedule
        else
        oa_damper_sch = model_add_schedule(model, oa_damper_sch)
        end

        # create a PSZ-AC for each zone
        air_loops = []
        thermal_zones.each do |zone|
        OpenStudio.logFree(OpenStudio::Info, 'openstudio.Model.Model', "Adding PSZ-AC for #{zone.name}.")

        air_loop = OpenStudio::Model::AirLoopHVAC.new(model)
        if system_name.nil?
        air_loop.setName("#{zone.name} PSZ-AC")
        else
        air_loop.setName("#{zone.name} #{system_name}")
        end

        # default design temperatures and settings used across all air loops
        dsgn_temps = std.standard_design_sizing_temperatures
        unless hot_water_loop.nil?
        hw_temp_c = hot_water_loop.sizingPlant.designLoopExitTemperature
        hw_delta_t_k = hot_water_loop.sizingPlant.loopDesignTemperatureDifference
        end

        # adjusted design heating temperature for psz_ac
        dsgn_temps['zn_htg_dsgn_sup_air_temp_f'] = 122.0
        dsgn_temps['zn_htg_dsgn_sup_air_temp_c'] = OpenStudio.convert(dsgn_temps['zn_htg_dsgn_sup_air_temp_f'], 'F', 'C').get
        dsgn_temps['htg_dsgn_sup_air_temp_f'] = dsgn_temps['zn_htg_dsgn_sup_air_temp_f']
        dsgn_temps['htg_dsgn_sup_air_temp_c'] = dsgn_temps['zn_htg_dsgn_sup_air_temp_c']

        # default design settings used across all air loops
        sizing_system = std.adjust_sizing_system(air_loop, dsgn_temps, min_sys_airflow_ratio: 1.0)

        # air handler controls
        # add a setpoint manager single zone reheat to control the supply air temperature
        setpoint_mgr_single_zone_reheat = OpenStudio::Model::SetpointManagerSingleZoneReheat.new(model)
        setpoint_mgr_single_zone_reheat.setName("#{zone.name} Setpoint Manager SZ Reheat")
        setpoint_mgr_single_zone_reheat.setControlZone(zone)
        setpoint_mgr_single_zone_reheat.setMinimumSupplyAirTemperature(dsgn_temps['zn_clg_dsgn_sup_air_temp_c'])
        setpoint_mgr_single_zone_reheat.setMaximumSupplyAirTemperature(dsgn_temps['zn_htg_dsgn_sup_air_temp_c'])
        setpoint_mgr_single_zone_reheat.addToNode(air_loop.supplyOutletNode)

        # zone sizing
        sizing_zone = zone.sizingZone
        sizing_zone.setZoneCoolingDesignSupplyAirTemperature(dsgn_temps['zn_clg_dsgn_sup_air_temp_c'])
        sizing_zone.setZoneHeatingDesignSupplyAirTemperature(dsgn_temps['zn_htg_dsgn_sup_air_temp_c'])

        # create heating coil
        case heating_type
        when 'NaturalGas', 'Gas'
        htg_coil = std.create_coil_heating_gas(model,
                                    name: "#{air_loop.name} Gas Htg Coil")
        when 'Water'
        if hot_water_loop.nil?
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.model.Model', 'No hot water plant loop supplied')
        return false
        end
        htg_coil = std.create_coil_heating_water(model,
                                    hot_water_loop,
                                    name: "#{air_loop.name} Water Htg Coil",
                                    rated_inlet_water_temperature: hw_temp_c,
                                    rated_outlet_water_temperature: (hw_temp_c - hw_delta_t_k),
                                    rated_inlet_air_temperature: dsgn_temps['prehtg_dsgn_sup_air_temp_c'],
                                    rated_outlet_air_temperature: dsgn_temps['htg_dsgn_sup_air_temp_c'])
        when 'Single Speed Heat Pump'
        htg_coil = std.create_coil_heating_dx_single_speed(model,
                                                name: "#{zone.name} HP Htg Coil",
                                                type: 'PSZ-AC',
                                                cop: 3.3)
        when 'Water To Air Heat Pump'
        htg_coil = std.create_coil_heating_water_to_air_heat_pump_equation_fit(model,
                                                                    hot_water_loop,
                                                                    name: "#{air_loop.name} Water-to-Air HP Htg Coil")
        when 'Electricity', 'Electric'
        htg_coil = std.create_coil_heating_electric(model,
                                        name: "#{air_loop.name} Electric Htg Coil")
        else
        # zero-capacity, always-off electric heating coil
        htg_coil = std.create_coil_heating_electric(model,
                                        name: "#{air_loop.name} No Heat",
                                        schedule: model.alwaysOffDiscreteSchedule,
                                        nominal_capacity: 0.0)
        end

        # create supplemental heating coil
        case supplemental_heating_type
        when 'Electricity', 'Electric'
        supplemental_htg_coil = std.create_coil_heating_electric(model,
                                                    name: "#{air_loop.name} Electric Backup Htg Coil")
        when 'NaturalGas', 'Gas'
        supplemental_htg_coil = std.create_coil_heating_gas(model,
                                                name: "#{air_loop.name} Gas Backup Htg Coil")
        else
        # Zero-capacity, always-off electric heating coil
        supplemental_htg_coil = std.create_coil_heating_electric(model,
                                                    name: "#{air_loop.name} No Heat",
                                                    schedule: model.alwaysOffDiscreteSchedule,
                                                    nominal_capacity: 0.0)
        end

        # create cooling coil
        case cooling_type
        when 'Water'
        if chilled_water_loop.nil?
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.model.Model', 'No chilled water plant loop supplied')
        return false
        end
        clg_coil = std.create_coil_cooling_water(model,
                                    chilled_water_loop,
                                    name: "#{air_loop.name} Water Clg Coil")
        when 'Two Speed DX AC'
        clg_coil = std.create_coil_cooling_dx_two_speed(model,
                                            name: "#{air_loop.name} 2spd DX AC Clg Coil")
        when 'Single Speed DX AC'
        clg_coil = std.create_coil_cooling_dx_single_speed(model,
                                                name: "#{air_loop.name} 1spd DX AC Clg Coil",
                                                type: 'PSZ-AC')
        when 'Single Speed Heat Pump'
        clg_coil = std.create_coil_cooling_dx_single_speed(model,
                                                name: "#{air_loop.name} 1spd DX HP Clg Coil",
                                                type: 'Heat Pump')
        # clg_coil.setMaximumOutdoorDryBulbTemperatureForCrankcaseHeaterOperation(OpenStudio::OptionalDouble.new(10.0))
        # clg_coil.setRatedSensibleHeatRatio(0.69)
        # clg_coil.setBasinHeaterCapacity(10)
        # clg_coil.setBasinHeaterSetpointTemperature(2.0)
        when 'Water To Air Heat Pump'
        if chilled_water_loop.nil?
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.model.Model', 'No chilled water plant loop supplied')
        return false
        end
        clg_coil = std.create_coil_cooling_water_to_air_heat_pump_equation_fit(model,
                                                                    chilled_water_loop,
                                                                    name: "#{air_loop.name} Water-to-Air HP Clg Coil")
        else
        clg_coil = nil
        end

        # Use a Fan:OnOff in the unitary system object
        case fan_type
        when 'Cycling'
        fan = std.create_fan_by_name(model,
                        'Packaged_RTU_SZ_AC_Cycling_Fan',
                        fan_name: "#{air_loop.name} Fan")
        when 'ConstantVolume'
        fan = std.create_fan_by_name(model,
                        'Packaged_RTU_SZ_AC_CAV_OnOff_Fan',
                        fan_name: "#{air_loop.name} Fan")
        else
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.model.Model', 'Invalid fan_type')
        return false
        end

        # fan location
        if fan_location.nil?
        fan_location = 'DrawThrough'
        end
        case fan_location
        when 'DrawThrough', 'BlowThrough'
        OpenStudio.logFree(OpenStudio::Debug, 'openstudio.Model.Model', "Setting fan location for #{fan.name} to #{fan_location}.")
        else
        OpenStudio.logFree(OpenStudio::Error, 'openstudio.model.Model', "Invalid fan_location #{fan_location} for fan #{fan.name}.")
        return false
        end

        # construct unitary system object
        unitary_system = OpenStudio::Model::AirLoopHVACUnitarySystem.new(model)
        unitary_system.setSupplyFan(fan) unless fan.nil?
        unitary_system.setHeatingCoil(htg_coil) unless htg_coil.nil?
        unitary_system.setCoolingCoil(clg_coil) unless clg_coil.nil?
        unitary_system.setSupplementalHeatingCoil(supplemental_htg_coil) unless supplemental_htg_coil.nil?
        unitary_system.setControllingZoneorThermostatLocation(zone)
        unitary_system.setFanPlacement(fan_location)
        unitary_system.addToNode(air_loop.supplyInletNode)

        # added logic and naming for heat pumps
        case heating_type
        when 'Water To Air Heat Pump'
        unitary_system.setMaximumOutdoorDryBulbTemperatureforSupplementalHeaterOperation(OpenStudio.convert(40.0, 'F', 'C').get)
        unitary_system.setName("#{air_loop.name} Unitary HP")
        unitary_system.setMaximumSupplyAirTemperature(dsgn_temps['zn_htg_dsgn_sup_air_temp_c'])
        unitary_system.setSupplyAirFlowRateMethodDuringCoolingOperation('SupplyAirFlowRate')
        unitary_system.setSupplyAirFlowRateMethodDuringHeatingOperation('SupplyAirFlowRate')
        unitary_system.setSupplyAirFlowRateMethodWhenNoCoolingorHeatingisRequired('SupplyAirFlowRate')
        when 'Single Speed Heat Pump'
        unitary_system.setMaximumOutdoorDryBulbTemperatureforSupplementalHeaterOperation(OpenStudio.convert(40.0, 'F', 'C').get)
        unitary_system.setName("#{air_loop.name} Unitary HP")
        else
        unitary_system.setName("#{air_loop.name} Unitary AC")
        end

        # specify control logic
        unitary_system.setAvailabilitySchedule(hvac_op_sch)
        if fan_type == 'Cycling'
        unitary_system.setSupplyAirFanOperatingModeSchedule(model.alwaysOffDiscreteSchedule)
        else # constant volume operation
        unitary_system.setSupplyAirFanOperatingModeSchedule(hvac_op_sch)
        end

        # add the OA system
        #oa_controller = OpenStudio::Model::ControllerOutdoorAir.new(model)
        #oa_controller.setName("#{air_loop.name} OA System Controller")
        #oa_controller.setMinimumOutdoorAirSchedule(oa_damper_sch)
        #oa_controller.autosizeMinimumOutdoorAirFlowRate
        #oa_controller.resetEconomizerMinimumLimitDryBulbTemperature
        #oa_system = OpenStudio::Model::AirLoopHVACOutdoorAirSystem.new(model, oa_controller)
        #oa_system.setName("#{air_loop.name} OA System")
        #oa_system.addToNode(air_loop.supplyInletNode)

        # TODO: enable economizer maximum fraction outdoor air schedule input
        # econ_eff_sch = model_add_schedule(model, 'RetailStandalone PSZ_Econ_MaxOAFrac_Sch')

        # set air loop availability controls and night cycle manager, after oa system added
        air_loop.setAvailabilitySchedule(hvac_op_sch)
        air_loop.setNightCycleControlType('CycleOnAny')
        avail_mgr = air_loop.availabilityManager
        if avail_mgr.is_initialized
        avail_mgr = avail_mgr.get
        if avail_mgr.to_AvailabilityManagerNightCycle.is_initialized
        avail_mgr = avail_mgr.to_AvailabilityManagerNightCycle.get
        avail_mgr.setCyclingRunTime(1800)
        end
        end

        # create a diffuser and attach the zone/diffuser pair to the air loop
        diffuser = OpenStudio::Model::AirTerminalSingleDuctUncontrolled.new(model, model.alwaysOnDiscreteSchedule)
        diffuser.setName("#{air_loop.name} Diffuser")
        air_loop.multiAddBranchForZone(zone, diffuser.to_HVACComponent.get)
        air_loops << air_loop
        end

        return air_loops

    end


    # Creates a dx dehmidifier for each zone and adds it to the model.
    #
    # @param thermal_zones [Array<OpenStudio::Model::ThermalZone>] array of zones to connect to this system
    def model_add_zone_dehumidifier_dx(model,
                                        thermal_zones,
                                        rated_water_removal_lday: 353,
                                        rated_energy_factor: 3.5,
                                        rated_airflow_cfm: 1750)

        dehumidifiers = []

        thermal_zones.each do |zone|
            
            dehumidifier = OpenStudio::Model::ZoneHVACDehumidifierDX.new(model)
            # set the dehmidifier properties 
            dehumidifier.setName("#{zone.name} dehumidifier")
            
            # rated_water_removal_lday = 353 # water removal in liters per day [L/day] - this is roughly 93.25 gallons/day
            # rated_energy_factor = 3.5 # liters per kWh [L/kWh]
            # rated_airflow_cfm = 1750

            rated_airflow_m3s = OpenStudio.convert(rated_airflow_cfm, 'cfm', 'm^3/s').get

            dehumidifier.setRatedWaterRemoval(rated_water_removal_lday)
            dehumidifier.setRatedEnergyFactor(rated_energy_factor)
            dehumidifier.setRatedAirFlowRate(rated_airflow_m3s)
            
            dehumidifier.addToThermalZone(zone)
            dehumidifiers << dehumidifier 
        end 

        return dehumidifiers 
    end


    # Creates a central heat pump for hot water chilled water and condenser water loops 
    #
    # @param 
    def model_add_central_heat_pump(model,
                                    system_name:"heat pump",
                                    hot_water_loop: nil,
                                    chilled_water_loop: nil,
                                    condenser_water_loop: nil)

        heat_pump = OpenStudio::Model::CentralHeatPumpSystem.new(model)
        # set the heat pump properties 
        heat_pump.setName(system_name)

        testo = hot_water_loop.addSupplyBranchForComponent(heat_pump)
        puts testo
        testo = chilled_water_loop.addSupplyBranchForComponent(heat_pump)
        puts testo
        condenser_water_loop.addSupplyBranchForComponent(heat_pump)

        return heat_pump 
    end


    
end