#import the pandapower module
import pandapower as pp
import numpy as np
import pandas as pd
import time
import matplotlib as plt
from openpyxl import load_workbook

import matplotlib.pyplot as plt
import matplotlib as mpl
import datetime as dt
from matplotlib.animation import FuncAnimation
from plotly import graph_objects as go
import plotly.express as px

numofgenerators = 19
data_gen = pd.read_csv('data_gen.csv')
precdf = pd.read_excel("Prec_Generation1.xlsx")
storagedata = {'Storage':[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]}
storagedf = pd.DataFrame(storagedata)
counter=0

consumpprecent = 0.47 ## assume 33.33% of the consumption
HouseholdNorth = consumpprecent*data_gen['household']
CommercialNorth = consumpprecent*data_gen['commercial']
IndustrialNorth = consumpprecent*data_gen['factories']

centregen=0

df1 = pd.DataFrame(index=[],columns=['date', 'generation', 'consumption'])

with open("System_Output.txt", "a") as f:

    for indx in range(48):
        sum_consump_nort = consumpprecent * data_gen['household'][indx] + consumpprecent * data_gen['commercial'][
            indx] + consumpprecent * data_gen['commercial'][indx]

        precdf = pd.read_excel("Prec_Generation1.xlsx")
        print("Date and Time:", data_gen["dates"][indx], file=f)

        for i in range(19):
            if (precdf["in_service_centre"][i] == True):
                if (precdf["Stations"][i] == "Rutenberg Power Staion"):
                    centregen = precdf["precentage"][i] * data_gen['coalgen'][indx] ##study case is if we add +centregen
                    break
                else:
                    centregen = precdf["precentage"][i] * data_gen['gasgen'][indx] ##study case is if we add +centregen
                    break
        mindisttonorth = precdf["distance_to_north_vpp"].min()
        print("Minimum distination to closest VPP:", mindisttonorth, file=f)
        mindistdf = precdf["distance_to_north_vpp"].nsmallest(2)
        print("Second Nearest to the same VPP:",mindistdf, file=f)
        nearstgenindx = precdf["distance_to_north_vpp"].idxmin()
        print("The nearst generator is:", precdf["Stations"][nearstgenindx], file=f)

        HouseholdCentre = 0.333 * data_gen['household'][indx]
        CommercialCentre = 0.333 * data_gen['commercial'][indx]
        IndustrialCentre = 0.333 * data_gen['factories'][indx]
        sumconsumpcentre = HouseholdCentre + CommercialCentre + IndustrialCentre
        flag_centre=0
        if(centregen > sumconsumpcentre ):
            flag_centre=1

        # print('gen',centregen)
        # print('sumconsump',sumconsumpcentre)

        #create an empty network
        net = pp.create_empty_network()

        #mainbus EHV 380KV TL
        pp.create_bus(net, name='MainBusNorth0', vn_kv=380, type='b')
        pp.create_bus(net, name='MainBusNorth', vn_kv=380, type='b')

        hv0_bus = pp.get_element_index(net, "bus", "MainBusNorth0")
        hv_bus = pp.get_element_index(net, "bus", "MainBusNorth")

        pp.create_line_from_parameters(net,from_bus=hv0_bus,to_bus=hv_bus,length_km=0.5,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='TL380KV')

        ##Generation
        pp.create_bus(net, name='AlonTavorPS', vn_kv=380, type='b')
        pp.create_bus(net, name='HagitPS', vn_kv=380, type='b')
        pp.create_bus(net, name='HaifaPS', vn_kv=380, type='b')
        pp.create_bus(net, name='OPCRotemPS', vn_kv=380, type='b')
        pp.create_bus(net, name='OrotRabinPS', vn_kv=380, type='b')
        pp.create_bus(net, name='ReadingPS', vn_kv=380, type='b')
        pp.create_bus(net, name='EshkolPS', vn_kv=380, type='b')
        pp.create_bus(net, name='GezerPS', vn_kv=380, type='b')
        pp.create_bus(net, name='DalyaPS', vn_kv=380, type='b')
        pp.create_bus(net, name='RutenbergPS', vn_kv=380, type='b')
        pp.create_bus(net, name='DoradPS', vn_kv=380, type='b')
        pp.create_bus(net, name='RamatHovavPS', vn_kv=380, type='b')
        pp.create_bus(net, name='NegevSPS', vn_kv=380, type='b')
        pp.create_bus(net, name='RamatHovavSPS', vn_kv=380, type='b')
        pp.create_bus(net, name='AshalimPlotA', vn_kv=380, type='b')
        pp.create_bus(net, name='AshalimPlotB', vn_kv=380, type='b')
        pp.create_bus(net, name='AshalimPlotC', vn_kv=380, type='b')


        ##add lines
        AlonTavorPS_bus = pp.get_element_index(net, "bus", "AlonTavorPS")
        HagitPS_bus = pp.get_element_index(net, "bus", "HagitPS")
        HaifaPS_bus = pp.get_element_index(net, "bus", "HaifaPS")
        OPCRotemPS_bus = pp.get_element_index(net, "bus", "OPCRotemPS")
        OrotRabinPS_bus = pp.get_element_index(net, "bus", "OrotRabinPS")
        ReadingPS_bus = pp.get_element_index(net, "bus", "ReadingPS")
        EshkolPS_bus = pp.get_element_index(net, "bus", "EshkolPS")
        GezerPS_bus = pp.get_element_index(net, "bus", "GezerPS")
        DalyaPS_bus = pp.get_element_index(net, "bus", "DalyaPS")
        RutenbergPS_bus = pp.get_element_index(net, "bus", "RutenbergPS")
        DoradPS_bus = pp.get_element_index(net, "bus", "DoradPS")
        NegevSPS_bus = pp.get_element_index(net, "bus", "NegevSPS")
        RamatHovavSPS_bus = pp.get_element_index(net, "bus", "RamatHovavSPS")
        RamatHovavPS_bus = pp.get_element_index(net, "bus", "RamatHovavPS")
        AshalimPlotA_bus = pp.get_element_index(net, "bus", "AshalimPlotA")
        AshalimPlotB_bus = pp.get_element_index(net, "bus", "AshalimPlotB")
        AshalimPlotC_bus = pp.get_element_index(net, "bus", "AshalimPlotC")

        ##crate generators:
        pp.create_gen(net,RutenbergPS_bus,p_mw=precdf["precentage"][0]*data_gen['coalgen'][indx],name = 'RutenbergPS',type= 'CG',in_service=precdf["in_service_north"][0])
        pp.create_gen(net,OrotRabinPS_bus,p_mw=precdf["precentage"][1]*data_gen['coalgen'][indx],name = 'OrotRabinPS',type= 'CG',in_service=precdf["in_service_north"][1])
        pp.create_gen(net,EshkolPS_bus,p_mw=precdf["precentage"][2]*data_gen['gasgen'][indx],name = 'EshkolPS',type= 'GG',in_service=precdf["in_service_north"][2])
        pp.create_gen(net,ReadingPS_bus,p_mw=precdf["precentage"][3]*data_gen['gasgen'][indx],name = 'ReadingPS',type= 'GG',in_service=precdf["in_service_north"][3])
        pp.create_gen(net,HaifaPS_bus,p_mw=precdf["precentage"][4]*data_gen['gasgen'][indx],name = 'HaifaPS',type= 'GG',in_service=precdf["in_service_north"][4])
        pp.create_gen(net,AlonTavorPS_bus,p_mw=precdf["precentage"][5]*data_gen['gasgen'][indx],name = 'AlonTavorPS',type= 'GG',in_service=precdf["in_service_north"][5])
        pp.create_gen(net,GezerPS_bus,p_mw=precdf["precentage"][6]*data_gen['gasgen'][indx],name = 'GezerPS',type= 'GG',in_service=precdf["in_service_north"][6])
        pp.create_gen(net,HagitPS_bus,p_mw=precdf["precentage"][7]*data_gen['gasgen'][indx],name = 'HagitPS',type= 'GG',in_service=precdf["in_service_north"][7])
        pp.create_gen(net,RamatHovavPS_bus,p_mw=precdf["precentage"][8]*data_gen['gasgen'][indx],name = 'RamatHovavPS',type= 'GG',in_service=precdf["in_service_north"][8])
        pp.create_gen(net,DoradPS_bus,p_mw=precdf["precentage"][9]*data_gen['gasgen'][indx],name = 'DoradPS',type= 'GG',in_service=precdf["in_service_north"][9])
        pp.create_gen(net,DalyaPS_bus,p_mw=precdf["precentage"][10]*data_gen['gasgen'][indx],name = 'DalyaPS',type= 'GG',in_service=precdf["in_service_north"][10])
        pp.create_gen(net,OPCRotemPS_bus,p_mw=precdf["precentage"][11]*data_gen['gasgen'][indx],name = 'OPCRotemPS',type= 'GG',in_service=precdf["in_service_north"][11])
        pp.create_gen(net,RamatHovavSPS_bus,p_mw=precdf["precentage"][12]*data_gen['solargen'][indx],name = 'RamatHovavSPS',type= 'SG',in_service=precdf["in_service_north"][12])
        pp.create_gen(net,NegevSPS_bus,p_mw=precdf["precentage"][13]*data_gen['solargen'][indx],name = 'NegevSPS',type= 'SG',in_service=precdf["in_service_north"][13])
        pp.create_gen(net,AshalimPlotA_bus,p_mw=precdf["precentage"][14]*data_gen['solargen'][indx],name = 'AshalimPlotA',type= 'SG',in_service=precdf["in_service_north"][14])
        pp.create_gen(net,AshalimPlotB_bus,p_mw=precdf["precentage"][15]*data_gen['solargen'][indx],name = 'AshalimPlotB',type= 'SG',in_service=precdf["in_service_north"][15])
        pp.create_gen(net,AshalimPlotC_bus,p_mw=precdf["precentage"][16]*data_gen['solargen'][indx],name = 'AshalimPlotC',type= 'SG',in_service=precdf["in_service_north"][16])



        print("Generators in service:",precdf["Stations"],net.gen.in_service, file=f)
        ##connect buses
        pp.create_line_from_parameters(net,from_bus=AlonTavorPS_bus,to_bus=hv0_bus,length_km=33.1,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='AlonTavorPS')
        pp.create_line_from_parameters(net,from_bus=HagitPS_bus,to_bus=hv0_bus,length_km=40,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='HagitPS')
        pp.create_line_from_parameters(net,from_bus=HaifaPS_bus,to_bus=hv0_bus,length_km=26.5,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='HaifaPS')
        pp.create_line_from_parameters(net,from_bus=OPCRotemPS_bus,to_bus=hv0_bus,length_km=61.4,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='OPCRotemPS')
        pp.create_line_from_parameters(net,from_bus=OrotRabinPS_bus,to_bus=hv0_bus,length_km=61.3,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='OrotRabinPS')
        pp.create_line_from_parameters(net,from_bus=ReadingPS_bus,to_bus=hv0_bus,length_km=102,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='ReadingPS')
        pp.create_line_from_parameters(net,from_bus=EshkolPS_bus,to_bus=hv0_bus,length_km=133.1,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='EshkolPS')
        pp.create_line_from_parameters(net,from_bus=DalyaPS_bus,to_bus=hv0_bus,length_km=133.1,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='DalyaPS')
        pp.create_line_from_parameters(net,from_bus=GezerPS_bus,to_bus=hv0_bus,length_km=136,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='GezerPS')
        pp.create_line_from_parameters(net,from_bus=RutenbergPS_bus,to_bus=hv0_bus,length_km=160.5,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='RutenbergPS')
        pp.create_line_from_parameters(net,from_bus=DoradPS_bus,to_bus=hv0_bus,length_km=159,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='DoradPS')
        pp.create_line_from_parameters(net,from_bus=NegevSPS_bus,to_bus=hv0_bus,length_km=204,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='NegevSPS')
        pp.create_line_from_parameters(net,from_bus=RamatHovavSPS_bus,to_bus=hv0_bus,length_km=203,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='RamatHovavSPS')
        pp.create_line_from_parameters(net,from_bus=RamatHovavPS_bus,to_bus=hv0_bus,length_km=202.5,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='RamatHovavPS')
        pp.create_line_from_parameters(net,from_bus=AshalimPlotA_bus,to_bus=hv0_bus,length_km=223,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='AshalimPlotA')
        pp.create_line_from_parameters(net,from_bus=AshalimPlotB_bus,to_bus=hv0_bus,length_km=223.5,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='AshalimPlotB')
        pp.create_line_from_parameters(net,from_bus=AshalimPlotC_bus,to_bus=hv0_bus,length_km=223.4,r_ohm_per_km=0.059,x_ohm_per_km=0.25,c_nf_per_km=14.6, max_i_ka=1.15,name='AshalimPlotC')


        ##add switches
        pp.create_switch(net,RutenbergPS_bus,pp.get_element_index(net,"line", "RutenbergPS"), et='l', closed=precdf["switch_north"][0], type='LBS', name='RutenbergPSSW')
        pp.create_switch(net,OrotRabinPS_bus,pp.get_element_index(net,"line", "OrotRabinPS"), et='l', closed=precdf["switch_north"][1], type='LBS', name='OrotRabinPSSW')
        pp.create_switch(net,EshkolPS_bus,pp.get_element_index(net,"line", "EshkolPS"), et='l', closed=precdf["switch_north"][2], type='LBS', name='EshkolPSSW')
        pp.create_switch(net,ReadingPS_bus,pp.get_element_index(net,"line", "ReadingPS"), et='l', closed=precdf["switch_north"][3], type='LBS', name='ReadingPSSW')
        pp.create_switch(net,HaifaPS_bus,pp.get_element_index(net,"line", "HaifaPS"), et='l', closed=precdf["switch_north"][4], type='LBS', name='HaifaPSSW')
        pp.create_switch(net,AlonTavorPS_bus,pp.get_element_index(net,"line", "AlonTavorPS"), et='l', closed=precdf["switch_north"][5], type='LBS', name='AlonTavorSW')
        pp.create_switch(net,GezerPS_bus,pp.get_element_index(net,"line", "GezerPS"), et='l', closed=precdf["switch_north"][6], type='LBS', name='GezerPSSW')
        pp.create_switch(net,HagitPS_bus,pp.get_element_index(net,"line", "HagitPS"), et='l', closed=precdf["switch_north"][7], type='LBS', name='HagitSW')
        pp.create_switch(net,RamatHovavPS_bus,pp.get_element_index(net,"line", "RamatHovavPS"), et='l', closed=precdf["switch_north"][8], type='LBS', name='RamatHovavPSSW')
        pp.create_switch(net,DoradPS_bus,pp.get_element_index(net,"line", "DoradPS"), et='l', closed=precdf["switch_north"][9], type='LBS', name='DoradPSSW')
        pp.create_switch(net,DalyaPS_bus,pp.get_element_index(net,"line", "DalyaPS"), et='l', closed=precdf["switch_north"][10], type='LBS', name='DalyaPSSW')
        pp.create_switch(net,OPCRotemPS_bus,pp.get_element_index(net,"line", "OPCRotemPS"), et='l', closed=precdf["switch_north"][11], type='LBS', name='OPCRotemPSSW')
        pp.create_switch(net,RamatHovavSPS_bus,pp.get_element_index(net,"line", "RamatHovavSPS"), et='l', closed=precdf["switch_north"][12], type='LBS', name='RamatHovavSPSSW')
        pp.create_switch(net,NegevSPS_bus,pp.get_element_index(net,"line", "NegevSPS"), et='l', closed=precdf["switch_north"][13], type='LBS', name='NegevSPSSW')
        pp.create_switch(net,AshalimPlotA_bus,pp.get_element_index(net,"line", "AshalimPlotA"), et='l', closed=precdf["switch_north"][14], type='LBS', name='AshalimPlotASW')
        pp.create_switch(net,AshalimPlotB_bus,pp.get_element_index(net,"line", "AshalimPlotB"), et='l', closed=precdf["switch_north"][15], type='LBS', name='AshalimPlotBSW')
        pp.create_switch(net,AshalimPlotC_bus,pp.get_element_index(net,"line", "AshalimPlotC"), et='l', closed=precdf["switch_north"][16], type='LBS', name='AshalimPlotCSW')

        # print(net.switch.closed)
        #HV Bus 110KV
        pp.create_bus(net, name='HVBusNorth0', vn_kv=110, type='b')
        lv0_bus = pp.get_element_index(net, "bus", "HVBusNorth0")

        pp.create_bus(net, name='HVBusNorth', vn_kv=110, type='b')
        lv_bus = pp.get_element_index(net, "bus", "HVBusNorth")

        #connect EHV-HV
        pp.create_transformer_from_parameters(net, hv_bus, lv0_bus, sn_mva=300, vn_hv_kv=380, vn_lv_kv=110, vkr_percent=0.06,
                                              vk_percent=8, pfe_kw=0, i0_percent=0, tp_pos=0, shift_degree=0, name='EHV-HV-Trafo')
        #creating line

        pp.create_line_from_parameters(net,from_bus=lv0_bus,to_bus=lv_bus,length_km=0.5,r_ohm_per_km=.1,x_ohm_per_km=0.05,c_nf_per_km=10,max_i_ka=0.4)

        ## add wind parks to 100KVBus
        pp.create_gen(net,lv0_bus,p_mw=precdf["precentage"][17]*data_gen['windgen'][indx],name = 'SirinWPS',type= 'WP',in_service=precdf["in_service_north"][17])
        pp.create_gen(net,lv0_bus,p_mw=precdf["precentage"][18]*data_gen['windgen'][indx],name = 'GilboaWPS',type= 'WP',in_service=precdf["in_service_north"][18])


        pp.create_bus(net, name="Bus MV0 20K Industrial North", vn_kv=20, type='n')
        pp.create_bus(net, name="Bus MV0 10K Commercial North", vn_kv=10, type='n')
        pp.create_bus(net, name="Bus MV1 10K Commercial North", vn_kv=10, type='n')


        mv1_bus = pp.get_element_index(net, "bus", "Bus MV0 20K Industrial North")
        lv1_bus = pp.get_element_index(net, "bus", "Bus MV0 10K Commercial North")
        pp.create_transformer3w_from_parameters(net, lv_bus, mv1_bus, lv1_bus, vn_hv_kv=110, vn_mv_kv=20, vn_lv_kv=10,
                                                sn_hv_mva=40, sn_mv_mva=15, sn_lv_mva=25, vk_hv_percent=10.1,
                                                vk_mv_percent=10.1, vk_lv_percent=10.1, vkr_hv_percent=0.266667,
                                                vkr_mv_percent=0.033333, vkr_lv_percent=0.04, pfe_kw=0, i0_percent=0,
                                                shift_mv_degree=30, shift_lv_degree=30, tap_side="hv", tap_neutral=0, tap_min=-8,
                                                tap_max=8, tap_step_percent=1.25, tap_pos=0, name='HV-MV-MV-Trafo')


        pp.create_bus(net, name="Bus LV0 0.4K Residential North", vn_kv=0.4, type='n')

        lv2_bus = pp.get_element_index(net, "bus", "Bus MV1 10K Commercial North")
        lv3_bus = pp.get_element_index(net, "bus", "Bus LV0 0.4K Residential North")

        pp.create_line_from_parameters(net,from_bus=lv1_bus,to_bus=lv2_bus,length_km=0.5,r_ohm_per_km=.1,x_ohm_per_km=0.05,c_nf_per_km=10,max_i_ka=0.4)

        pp.create_transformer_from_parameters(net, lv2_bus, lv3_bus, sn_mva=0.4, vn_hv_kv=10, vn_lv_kv=0.4, vkr_percent=1.325, vk_percent=4, pfe_kw=0.95, i0_percent=0.2375, tap_side="hv", tap_neutral=0, tap_min=-2, tap_max=2, tap_step_percent=2.5, tp_pos=0, shift_degree=150, name='MV-LV-Trafo')

        ##create 1 Commercial load on 10KV Bus
        pp.create_load(net, lv2_bus, p_mw=CommercialNorth[indx], q_mvar=0, name="CommercialNorth")

        ##create 1 Industrial load on 20KV Bus
        pp.create_load(net, mv1_bus, p_mw=IndustrialNorth[indx], q_mvar=0, name="IndustrialNorth")

        ##create 1  bus and load on the 400v line
        pp.create_bus(net, name='Bus LV1.1', vn_kv=0.4, type='m')
        HouseholdNorth_idx = pp.get_element_index(net, "bus", "Bus LV1.1")

        pp.create_line_from_parameters(net,from_bus=lv3_bus,to_bus=HouseholdNorth_idx,length_km=1,r_ohm_per_km=0.161,x_ohm_per_km=0.117,c_nf_per_km=273, max_i_ka=0.4)
        pp.create_load(net, HouseholdNorth_idx, p_mw=HouseholdNorth[indx], q_mvar=0, name="HouseholdNorth")

        sum_gen = 0
        count=19
        i=0

        for i in range(count):
            if net.gen.in_service[i] == True:
                print(net.gen.name[i],net.gen.p_mw[i])
                sum_gen = net.gen.p_mw[i]+sum_gen;

        print("Sum of all generation:",sum_gen, file=f)
        # print(sum_gen-sum(net.load.p_mw))
        print("Net Load",net.load, file=f)
        storsum = sum(storagedf["Storage"])

        indx += 1
        overload = 0
        load = 0

    # VPP
        while sum_gen > sum(net.load.p_mw):
            max_e_mwh = 6000
            excessenergy = sum_gen - sum(net.load.p_mw)
            storagedf["Storage"][counter] =+ excessenergy
            print(storagedf["Storage"])
            storsum = sum(storagedf["Storage"]) + storsum
            counter = counter+1
            print("Available stored energy",storsum, file=f)

            if storsum < max_e_mwh:
                pp.create_storage(net, lv0_bus, storsum, max_e_mwh, name='north_storage')
                print(net.storage)
                break;
            else:
                pp.create_storage(net, lv0_bus, max_e_mwh, max_e_mwh, name='north_storage')
                print("Export Excess Energy:",excessenergy, "MWH", file=f)
                print(net.storage)
                break;
        else:
            overload = sum(net.load.p_mw)-sum_gen
            if (precdf["in_service_north"].equals(precdf["no_fault_in_service_north"])):
                print("no fault in service, only overload", file=f)
                print("overload is,", overload, ",reduce the load!", file=f)
                print("check for stored energy, stored energy is:", storsum, file=f)
                if(storsum>0):
                    if (overload < storsum):
                        storsum = storsum - overload
                        overload = 0
                        print("Stored energy is available and grater than the overload", file=f)
                        print("Overload has been reduced", file=f)
                    else:

                        newoverload = overload-storsum
                        storsum = 0
                        print("shutdown some part of the network, previous overload is,", overload, " new overload is",
                              newoverload, file=f)
                else:
                    print("there is no stored energy, shutdown some part of the network!", file=f)
                    load = sum_gen - 2 * overload
                    print("previous load is,", sum(net.load.p_mw), " new load is", load, file=f)
            else:
                for i in range(count):
                    if (precdf["in_service_north"][i] != precdf["no_fault_in_service_north"][i]):
                        print("Fault in service...fix " + precdf["Stations"][
                                i] + " generator", file=f)
                        print("Overload,", overload, ",check for stored energy!", file=f)
                        print("stored energy is :",storsum, file=f)
                        print("all stored energy is :", storsum, file=f)
                if(overload < storsum):
                    storsum = storsum-overload
                    overload = 0
                    print("No overload", file=f)
                    print("Current Stored Energy is",storsum, file=f)
                else:
                    if(flag_centre == 1):
                        if((nearstgenindx == 1 ) and ( precdf["precentage"][nearstgenindx] * data_gen['coalgen'][indx] > overload)):
                            updated_centregen = centregen - overload
                            print("overload is dealt with", file=f)
                            print("used energy from ",precdf["Stations"][nearstgenindx], 'is',overload, file=f)
                            print("the new generation in the center is",updated_centregen, file=f)
                            updated_centregen = 0
                        if(precdf["precentage"][nearstgenindx] * data_gen['gasgen'][indx] > overload):
                            updated_centregen = centregen - overload
                            print("overload is dealt with", file=f)
                            print("used energy from ",precdf["Stations"][nearstgenindx], 'is',overload, file=f)
                            print("the new generation in the center is",updated_centregen, file=f)
                            updated_centregen = 0
                        else:
                             for j in range(count):
                                if(precdf["distance_to_north_vpp"][j] == mindistdf[2]):
                                    print("the second nearest generator is:",precdf["Stations"][j], file=f)
                                    print("the used energy from this station is:",precdf["precentage"][j]*data_gen["gasgen"][indx], file=f)
                    else:
                        print("There is no available generators for help", file=f)
                        print("disconnect few loads", file=f)
                        print("Previous Load is",sum(net.load.p_mw),"New Load is:",sum(net.load.p_mw)-overload, file=f)
            print("", file=f)


        df2 = pd.DataFrame([[data_gen["dates"][indx], sum_gen, sum(net.load.p_mw), storsum, load, overload]], index=[indx],
                          columns=['date', 'generation', 'consumption', 'storedenergy','newload', 'newoverload'])
        df1 = df1.append(df2)
        print(df1)

        df1.to_excel("dataoutputpy.xlsx")

        time.sleep(0.001)

