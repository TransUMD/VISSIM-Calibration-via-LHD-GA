# VISSIM Parameter Calibration
# Latin Hypercube Sampling to reduce range
# Genetic Algorithm to find optimal parameters

# Author: M.Y.

import numpy as np
import matplotlib.pyplot as plt
import os
import sys
import optparse
import subprocess
import random
import time
import re
import win32com.client as com # VISSIM COM
import geatpy as ga           # Genetic Algorithm Package
import time
import xlrd
from pyDOE import *
from scipy.stats.distributions import norm

# Fitness Calculation for GA
def aimfuc(Phen, LegV, Run, NIND):
    f = np.zeros((NIND,1));
    #FS = np.array([Phen[:,0]]).T
    for q in range(NIND):
        C0 = np.array([Phen[:,0]]).T[q][0]
        C2 = np.array([Phen[:,1]]).T[q][0]
        #Obsr = np.array([Phen[:,1]]).T
        LHD = np.array([Phen[:,2]]).T[q][0]
        MHD = np.array([Phen[:,3]]).T[q][0]
        ESD = np.array([Phen[:,4]]).T[q][0]
        LC = np.array([Phen[:,5]]).T[q][0]
    
        dbps.SetAttValue('W99cc0',C0)
        dbps.SetAttValue('W99cc2',C2)
        dbps.SetAttValue('LookAheadDistMax',LHD)
        dbps.SetAttValue('MinHdwy',MHD)
        for i in range(len(Nlinks)):
            Nlinks[i].SetAttValue('EmergStopDist',ESD)
            Nlinks[i].SetAttValue('LnChgDist',LC)
        t1 = []; t2 = [];
        k = 10
        for j in range(0,TotalPeriod):
            Sim.RunSingleStep()
            if j>WarmPeriod and j%60 == 0:  # Start to collect data after 600 seconds
                k += 1
                t1.append(GTT1.AttValue('TravTm(Current,%i,All)'% k)) # 110-04626
                t2.append(GTT2.AttValue('TravTm(Current,%i,All)'% k)) # 110N046262
            if j==TotalPeriod-1:
                Sim.Stop()
        k += 1
        t1.append(GTT1.AttValue('TravTm(%i,%i,All)'% (Run,k))) # 110-04626
        t2.append(GTT2.AttValue('TravTm(%i,%i,All)'% (Run,k))) # 110N046262
        GT1.append(t1); GT2.append(t2);
        Run += 2
        
        fitness =  np.array(sum(abs(np.array(t1)-np.array(I1))+abs(np.array(t2)-np.array(I2))))
        GFit.append(fitness)
        print(fitness)
        f[q] = fitness
        
    return[f, LegV, Run]
    
if __name__ == "__main__":
#=================================================================================
    # Candidate parameter sets
    SFS = [90, 100, 120]                            # 3 Fixed           FS
    SC0 = [1.3, 1.4, 1.5, 1.6, 1.7]                 # 5 Not Fixed       C0
    SC1 = [1, 2]                                    # 2 Fixed           C1
    SC2 = [2, 3, 4, 5, 6, 7]                        # 6 Not Fixed       C2
    SObsr = [1, 2, 3, 4]                            # 4 Fixed           Observed Vehicles
    SLHD = [200, 225, 250, 275, 300]                # 5 Not Fixed       Look Ahead Distance
    SMHD = [0.45, 0.50, 0.55, 0.60]                 # 4 Not Fixed       Minimum Headway
    SESD = [5, 6, 7, 8]                             # 4 Not Fixed       Emergency Stop Distance
    SLC = [200, 300, 400, 500, 600]                 # 5 Not Fixed       Lane Change Distance
    Pairs = [SFS, SC0, SC1, SC2, SObsr, SLHD, SMHD, SESD, SLC]
    PairsLength = [len(SFS), len(SC0), len(SC1), len(SC2), len(SObsr), len(SLHD), len(SMHD), len(SESD), len(SLC)]
    Combinations = len(SFS)*len(SC0)*len(SC1)*len(SC2)*len(SObsr)*len(SLHD)*len(SMHD)*len(SESD)*len(SLC)
#=================================================================================
# Latin Hypercube Sampling Design
# Candidate sets
    sam = 300 # sample size
    LHsets = lhs(9, samples = sam)      				# Generate candidate sets            
    LHsets = norm(loc=0, scale=1).ppf(LHsets)       # Normalized the value of candidate sets to N(0,1)
    for i in range(9):                             # Substitute the LH matrix to real value
        Range = Pairs[i]
        N = PairsLength[i]
        Prob = 1/N
        Interval = [];
        for k in range(N-1):
            Interval.append(norm.ppf(Prob*(k+1)))
        for j in range(sam):
            for q in range(N-1):
                if Interval[q] > LHsets[j,i]:
                    LHsets[j,i] = Range[q]
                    break
                if q == N-2:
                    LHsets[j,i] = Range[N-1]
#=================================================================================
# Groundtruth Travel Time Data    
	# INRIX Travel Time Data
    Data = xlrd.open_workbook('Your Path');
    Table = Data.sheet_by_name(u'Sheet1')
    I1 = Table.col_values(0) # Travel time for segment 1
    I2 = Table.col_values(1) # Travel time for segment 2
#=================================================================================  
# VISSIM Configurations  
    # Load VISSIM Network
    Vissim = com.Dispatch("Vissim.Vissim.9");
    Vissim.LoadNet(r'Your VISSIM inpx file path');    
    
    # Define Simulation Configurations
    Sim = Vissim.Simulation;
    Net = Vissim.Net
    G = Vissim.Graphics
    dbpss = Net.DrivingBehaviors;                                               # Driving behavior module
    dbps = dbpss.ItemByKey(3);                                                  # Freeway - 3
    Rel_Flows = Net.VehicleCompositions.ItemByKey(1).VehCompRelFlows.GetAll()   # Vehicle composition All	
    TT1 = Net.VehicleTravelTimeMeasurements.ItemByKey(1)                        # Travel time collection for 110-04626
    TT2 = Net.VehicleTravelTimeMeasurements.ItemByKey(2)                        # Travel time collection for 110N04626
    lbpss = Net.Links                                                           # Link module
    Nlinks = lbpss.GetAll()                                                     # Get all links
    
    # Set Simulation Parameters
    TotalPeriod = 25200;        # Define total simulation period
    WarmPeriod = 600;           # Define warm period 10 minutes     
    Random_Seed = 42;           # Define random seed
    step_time = 1;              # Define Step Time
    Sim.SetAttValue('SimPeriod',TotalPeriod);
    Sim.SetAttValue('SimRes',step_time);
    Sim.SetAttValue('RandSeed', Random_Seed);   
    Sim.SetAttValue('NumRuns',1200)
    G.CurrentNetworkWindow.SetAttValue('QuickMode',1)
#=================================================================================
# Data Collection Variables
    # Travel time collection for each run
    t1 = []; t2 = [];
    # Travel time collection for all run
    T1 = []; T2 = [];
    # Fitness collection for all run
    Fit = [];
#=================================================================================
# Test all candidate sets, to reduce the boundary
    Run = 1     # Indexing for travel time collection convenience
    for i in range(sam):
        Rel_Flows[0].SetAttValue('DesSpeedDistr', LHsets[i,0])
        Rel_Flows[1].SetAttValue('DesSpeedDistr', LHsets[i,0])
        dbps.SetAttValue('W99cc0',LHsets[i,1])
        dbps.SetAttValue('W99cc1Distr',LHsets[i,2])
        dbps.SetAttValue('W99cc2',LHsets[i,3])
        dbps.SetAttValue('ObsrvdVehs',LHsets[i,4])
        dbps.SetAttValue('LookAheadDistMax',LHsets[i,5])
        dbps.SetAttValue('MinHdwy',LHsets[i,6])
        for i in range(len(Nlinks)):
            Nlinks[i].SetAttValue('EmergStopDist',LHsets[i,7])
            Nlinks[i].SetAttValue('LnChgDist',LHsets[i,8])
        # Each scenario run 5 times
        for z in range(1):
            t1 = []; t2 = [];
            k = 10
            for j in range(0,TotalPeriod):
                Sim.RunSingleStep()
                if j>WarmPeriod and j%60 == 0:  # Start to collect data after 600 seconds
                    k += 1
                    t1.append(TT1.AttValue('TravTm(Current,%i,All)'% k)) # 110-04626
                    t2.append(TT2.AttValue('TravTm(Current,%i,All)'% k)) # 110N046262
                if j==TotalPeriod-1:
                    Sim.Stop()
            k += 1
            t1.append(TT1.AttValue('TravTm(%i,%i,All)'% (Run,k))) # 110-04626
            t2.append(TT2.AttValue('TravTm(%i,%i,All)'% (Run,k))) # 110N046262
            T1.append(t1); T2.append(t2);
            Run += 2
            # Calculate the fitness function, cumulative travel time error
            Fitness = sum(abs(np.array(t1)-np.array(I1))+abs(np.array(t2)-np.array(I2)))
            print(Fitness)
            Fit.append(Fitness)
    Vissim = None;
#=================================================================================
# Process the data, preparing inputs for GA
    TravelTime = np.array(Fit);
    LHtt = np.c_[TravelTime, LHsets];
    LHtt = LHtt[LHtt[:,0].argsort()]
    CSets = LHtt[0:50,:];                               # Top 20 Candidate Sets
    #CSFS = [1, 3]                                       # 1=90, 2=100, 3=120
    CSC0 = [CSets[:,2].min(), CSets[:,2].max()]         # Not Fixed
    #CSC1 = [CSets[:,3].min(), CSets[:,3].max()]         # 1=0.5S, 2=0.9S
    CSC2 = [CSets[:,4].min(), CSets[:,4].max()]         # Not Fixed
    #CSObsr = [CSets[:,5].min(), CSets[:,5].max()]       # Fixed
    CSLHD = [CSets[:,6].min(), CSets[:,6].max()]        # Not Fixed
    CSMHD = [CSets[:,7].min(), CSets[:,7].max()]        # Not Fixed
    CSESD = [CSets[:,8].min(), CSets[:,8].max()]        # Not Fixed
    CSLC = [CSets[:,9].min(), CSets[:,9].max()]         # Not Fixed   
#===============================================================================
# Open a new VISSIM for GA
    # Load VISSIM network
    Vissim = com.Dispatch("Vissim.Vissim.9");
    Vissim.LoadNet(r'Your VISSIM inpx file path in another folder for GA');    
    
    # Define Simulation Configurations
    Sim = Vissim.Simulation;
    Net = Vissim.Net
    G = Vissim.Graphics
    dbpss = Net.DrivingBehaviors;                                               # Driving behavior module
    dbps = dbpss.ItemByKey(3);                                                  # Freeway - 3
    Rel_Flows = Net.VehicleCompositions.ItemByKey(1).VehCompRelFlows.GetAll()   # Vehicle composition All	
    GTT1 = Net.VehicleTravelTimeMeasurements.ItemByKey(1)                        # Travel time collection for 110-04626
    GTT2 = Net.VehicleTravelTimeMeasurements.ItemByKey(2)                        # Travel time collection for 110N04626
    lbpss = Net.Links                                                           # Link module
    Nlinks = lbpss.GetAll()                                                     # Get all links
    
    # Set Simulation Parameters, only the number of runs changed
    Sim.SetAttValue('SimPeriod',TotalPeriod);
    Sim.SetAttValue('SimRes',step_time);
    Sim.SetAttValue('RandSeed', Random_Seed);   
    Sim.SetAttValue('NumRuns',5000)
    G.CurrentNetworkWindow.SetAttValue('QuickMode',1)
    
    # Set the parameters don't need to be calibrated
    Rel_Flows[0].SetAttValue('DesSpeedDistr', 100)      # FS=100
    Rel_Flows[1].SetAttValue('DesSpeedDistr', 100)      # FS=100
    dbps.SetAttValue('W99cc1Distr',1)                   # CC1=0.5S
    dbps.SetAttValue('ObsrvdVehs',2)                    # Obsr=2vehicles
#===============================================================================
# Data Collection Variables
    # Travel time collection for each run
    t1 = []; t2 = [];
    # Travel time collection for all run
    GT1 = []; GT2 = [];
    # Fitness collection for all run
    GFit = [];
#=================================================================================
# Genetic Algorithm to find the optimal parameter value  
    Run = 1;    # Indexing for travel time collection convenience
    
    # GA preparation, please refer to Geatpy package for details
    b1 = [1, 1]; b2 = [1, 1]; b3 = [1, 1]; b4 = [1, 1]; b5 = [1, 1]; b6 = [1, 1]; # Boundary, 0 represent no boundary
    codes = [1, 1, 1, 1, 1, 1]                                                    # Coding method
    precisions = [2, 2, 0, 2, 2, 0]                                               # Precision, how many decimals
    scales = [0, 0, 0, 0, 0, 0]                                            
    ranges = np.vstack([CSC0, CSC2, CSLHD, CSMHD, CSESD, CSLC]).T      
    borders = np.vstack([b1, b2, b3, b4, b5, b6]).T      
    
    # GA Variables, please refer to Geatpy package for details
    NIND = 2                   
    MAXGEN = 100             
    GGAP = 0.8                  
    selectStyle = 'sus'         
    recombinStyle = 'xovdp'    
    recopt = 0.9                
    pm = 0.1                    
    SUBPOP = 1                  
    maxormin = 1               
    
    # Start of GA, please refer to Geatpy package for details
    FieldD = ga.crtfld(ranges, borders, precisions, codes, scales)  
    Lind = np.sum(FieldD[0, :])             
    Chrom = ga.crtbp(NIND, Lind)           
    Phen = ga.bs2rv(Chrom, FieldD)        
    LegV = np.ones((NIND, 1))            
    [ObjV, LegV, Run] = aimfuc(Phen, LegV, Run, NIND)      
    pop_trace = (np.zeros((MAXGEN,2)) * np.nan)
    ind_trace = (np.zeros((MAXGEN,Lind)) * np.nan)
    start_time = time.time()
    for gen in range(MAXGEN):
        FitnV = ga.ranking(maxormin * ObjV, LegV)
        SelCh = ga.selecting(selectStyle, Chrom, FitnV, GGAP, SUBPOP) 
        SelCh = ga.recombin(recombinStyle, SelCh, recopt, SUBPOP) 
        SelCh = ga.mutbin(SelCh, pm) 
        Phen = ga.bs2rv(SelCh, FieldD) 
        LegVSel = np.ones((SelCh.shape[0], 1)) 
        [ObjVSel, LegVSel, Run] = aimfuc(Phen, LegVSel, Run, NIND)
        [Chrom, ObjV, LegV] = ga.reins(Chrom, SelCh, SUBPOP, 1, 1, maxormin*ObjV, maxormin*ObjVSel, ObjV, ObjVSel, LegV, LegVSel)
        pop_trace[gen, 1] = np.sum(ObjV)/ObjV.shape[0]
        if maxormin == 1:
            best_ind = np.argmin(ObjV)
        elif maxormin == -1:
            best_ind = np.argmax(ObjV)
        pop_trace[gen, 0] = ObjV[best_ind]
        ind_trace[gen, :] = Chrom[best_ind, :]
    
    end_time = time.time()
    ga.trcplot(pop_trace,[['Best Individual Fitness', 'Population Average Fintess']], ['Result'])
    
#===============================================================================
# Output
    best_gen = np.argmin(pop_trace[:, 0])
    print('Optimal Fitness：', np.min(pop_trace[:, 0]))
    print('Optimal Variables：')
    variables = ga.bs2rv(ind_trace, FieldD)
    for i in range(variables.shape[1]):
        print(variables[best_gen,i])
    print('The best generation is', best_gen + 1, 'th generation')
    print('Time Spent：', end_time - start_time, 'seconds')