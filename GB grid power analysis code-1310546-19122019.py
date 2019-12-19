# -*- coding: utf-8 -*-
"""
Created on Tue Dec 17 10:42:01 2019

@author: erikm
"""

import pandas as pd
import numpy as np
import pandapower as pp
import matplotlib.pyplot as plt
import geopandas
import math
from pandapower.plotting.plotly import simple_plotly
import pandapower.networks as pn

net = pn.GBreducednetwork()

pq_load= ((net.load.p_mw)**2 + (net.load.q_mvar)**2)**0.5
pq_load_sum=pq_load.sum()
scaling_factor= 52000/pq_load_sum

print("Load in Case1888 =",int(pq_load_sum),"MW")
print("In 2018, peak load was 52,000 MW VAR")
print("To replicate this system loading and generation GBreducednetwork is multipled by ", scaling_factor ,"using scaling in Pandapower")

net.load.scaling=scaling_factor
net.gen.scaling=scaling_factor
net.sgen.scaling=scaling_factor


#pp.to_excel(net,'dftest.xlsx')

df1=pd.DataFrame()

for line in net.line.index:
    
    net = pn.GBreducednetwork()
#Scaling system to represnet worst-case loads
    net.load.scaling=scaling_factor
    net.gen.scaling=scaling_factor
    net.sgen.scaling=scaling_factor
    
    pp.drop_lines(net,lines=[line])
    

    #try:
    pp.runpp(net)
    #get the 3 lines with largest loading percent and list them in df1
    loading_percent = (net.res_line.loading_percent)
    loading_percent_max = loading_percent.nlargest(3)
    df1.at[line,'first_loading_percent_max'] = float(loading_percent_max.iloc[0])
    df1.at[line,'first_loading_percent_max_index'] = float(loading_percent_max.index[0])
    df1.at[line,'second_loading_percent_max'] = float(loading_percent_max.iloc[1])
    df1.at[line,'second_loading_percent_max_index'] = float(loading_percent_max.index[1])
    df1.at[line,'thrid_loading_percent_max'] = float(loading_percent_max.iloc[2])
    df1.at[line,'third_loading_percent_max_index'] = float(loading_percent_max.index[2])
    
    loading_percent_exceed_100_occurance = sum(i >100 for i in loading_percent)
    df1.at[line,'loading_percent_exceed_100_occurance'] = float(loading_percent_exceed_100_occurance)
    
    loading_percent = net.res_line.loading_percent
    loading_percent_ave = loading_percent.sum()/len(net.res_line.loading_percent)
    df1.at[line,'loading_percent_ave'] = float(loading_percent_ave)
    
    loading_percent = (net.res_line.loading_percent)
    loading_percent_max = loading_percent.max()
    df1.at[line,'loading_percent_max'] = float(loading_percent_max)
    
    loading_percent_max_index = loading_percent.idxmax()
    df1.at[line,'loading_percent_max_index'] = float(loading_percent_max_index)
    
    line_power = abs(net.res_line.p_from_mw)
    line_power_sum = line_power.sum()
    df1.at[line,'line_power_sum'] = float(line_power_sum)
    
    line_power = abs(net.res_line.p_from_mw)
    line_power_max = line_power.nlargest(3)
    df1.at[line,'first_line_power_max'] = float(line_power_max.iloc[0])
    df1.at[line,'first_line_power_index'] = float(line_power_max.index[0])
    df1.at[line,'second_line_power_max'] = float(line_power_max.iloc[1])
    df1.at[line,'second_line_power_max_index'] = float(line_power_max.index[1])
    df1.at[line,'thrid_line_power_max'] = float(line_power_max.iloc[2])
    df1.at[line,'third_line_power_max_index'] = float(line_power_max.index[2])
    
    volt_magto =(net.res_line.vm_to_pu)
    volt_magto_max = volt_magto.max()
    df1.at[line,'volt_magto_max'] = float(volt_magto_max)
    
    volt_magfrom =(net.res_line.vm_from_pu)
    volt_magfrom_max = volt_magfrom.max()
    df1.at[line,'volt_magfrom_max'] = float(volt_magfrom_max)
        

    #except:
    #error_variable = float("0")
    #df1.at[line,'loading_percent_ave'] = error_variable
    #df1.at[line,'loading_percent_max'] = error_variable  
    #df1.at[line,'line_power_sum'] = error_variable
    #df1.at[line,'line_power_max'] = error_variable
    #df1.at[line,'volt_magto_max'] = error_variable
    #df1.at[line,'volt_magfrom_max'] = error_variable
         


#df.to_csv('test3.csv')

#pp.to_excel(net,'scalingtest2.xlsx')
    
net = pn.GBreducednetwork()
net.load.scaling=scaling_factor
net.gen.scaling=scaling_factor
net.sgen.scaling=scaling_factor
df2=pd.DataFrame()

y = range(0,86)
for line in y:
    
    pp.drop_lines(net,lines=[line])

    try:
        pp.runpp(net)
        
        #Calculating the three largest bus voltages per unit, with corresponding indexes when each line is removed 
        bus_voltage = (net.res_bus.vm_pu)
        bus_voltage_max = bus_voltage.nlargest(3)
        df2.at[line,'first_bus_voltage_max'] = float(bus_voltage_max.iloc[0])
        df2.at[line,'first_bus_voltage_max_index'] = float(bus_voltage_max.index[0])
        df2.at[line,'second_bus_voltage_max'] = float(bus_voltage_max.iloc[1])
        df2.at[line,'second_bus_voltage_max_index'] = float(bus_voltage_max.index[0])
        df2.at[line,'third_bus_voltage_max'] = float(bus_voltage_max.iloc[2])
        df2.at[line,'third_bus_voltage_max_index'] = float(bus_voltage_max.index[2])
        
        #Calculating the three smallest bus voltages per unit, with corresponding indexes when each line is removed 
        bus_voltage_min = bus_voltage.nsmallest(3)
        df2.at[line,'first_bus_voltage_min'] = float(bus_voltage_min.iloc[0])
        df2.at[line,'first_bus_voltage_min_index'] = float(bus_voltage_min.index[0])
        df2.at[line,'second_bus_voltage_min'] = float(bus_voltage_min.iloc[1])
        df2.at[line,'second_bus_voltage_min_index'] = float(bus_voltage_min.index[0])
        df2.at[line,'third_bus_voltage_min'] = float(bus_voltage_min.iloc[2])
        df2.at[line,'third_bus_voltage_min_index'] = float(bus_voltage_min.index[2])
    
    except:
        #Error blocking ensures that the powerflow is completed. 
        #When the flow does not converge "0" is recorded and analysis is continued.
        error_variable = float("0")
        df2.at[line,'first_bus_voltage_max'] = error_variable
        df2.at[line,'first_bus_voltage_max_index'] = error_variable
        df2.at[line,'second_bus_voltage_max'] = error_variable
        df2.at[line,'second_bus_voltage_max_index'] = error_variable
        df2.at[line,'third_bus_voltage_max'] = error_variable
        df2.at[line,'third_bus_voltage_max_index'] = error_variable
        
        df2.at[line,'first_bus_voltage_min'] = error_variable
        df2.at[line,'first_bus_voltage_min_index'] = error_variable
        df2.at[line,'second_bus_voltage_min'] = error_variable
        df2.at[line,'second_bus_voltage_min_index'] = error_variable
        df2.at[line,'third_bus_voltage_min'] = error_variable
        df2.at[line,'third_bus_voltage_min_index'] = error_variable

#Saving bus results to csv for further analysis        
df2.to_csv('UK_Network_Bus_Results.csv')

df1