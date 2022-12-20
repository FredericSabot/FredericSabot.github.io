import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2019 SP4\Python\3.7")
import powerfactory as pf
app=pf.GetApplication()
from math import sqrt

if app is None:
     raise Exception("getting PowerFactory application failed")

user=app.GetCurrentUser()

IEEE9 = 0
if IEEE9:
     project_name = "Nine-bus System"
#     project_name = "Nine-bus System-Protections"
else:
     project_name = "39 Bus New England System"

project=user.GetContents(project_name + '.IntPrj')[0]
project.Activate()

lines = app.GetCalcRelevantObjects("*.ElmLne")

def setRelayVT(relay_VT):
    V = relay_VT.cn_bus.uknom # Nominal voltage of the substation in kV
    relay_VT.ptapset = V*1000
    if relay_VT.ptapset != V*1000:
        print("Bus voltage: " + str(V))
        raise NotImplementedError("Please add first add an appropriate winding to the VT type \"Voltage Transformer\" in local library")
    relay_VT.stapset = 110

"""
Transducers
"""
library = app.GetGlobalLibrary("TypRelay")
# distance_polygonal = library.GetContents("ProtGeneric")[0].GetContents("F21 Distance Polygonal")[0].GetContents("F21 Distance Polygonal")[0]
try:
    distance_polygonal = app.GetLocalLibrary().GetContents("F21 Distance Polygonal")[0]
except IndexError:
    raise NotImplementedError("Please copy the F21 relay from GlobalLibrary/Generic and then change the output Logic settings to be the same as in https://www.digsilent.de/en/faq-reader-powerfactory/is-it-possible-to-simulate-distance-protection-scheme-in-powerfactory.html")
CT = app.GetGlobalLibrary().GetContents("Types")[0].GetContents("Ct")[0].GetContents("CT 120-1000/1A.TypCt")[0]
interlink = library.GetContents("ProtGeneric")[0].GetContents("Interlink")[0].GetContents("Interlink")[0]
try:
    VT = app.GetLocalLibrary().GetContents("Voltage Transformer.TypVt")[0] # This needs to be created first and should allow a V/110 ratio where V is the voltage on any bus that has an over/underspeed or underfrequency load shedding relay
except IndexError:
    raise NotImplementedError("Please add a VT type in the local library with the name \"Voltage Transformer\", this VT will be used for all measurements and should have windings that allow to bring any used nominal voltage (e.g.: 345000, 23000, 16500) to 110V")


"""
Distance Protection
"""
for line in lines:
    for side in range(2):
        cub = line.GetCubicle(side)
        old_relay = cub.cpRelays
        while old_relay != None:
            old_relay.Delete()
            old_relay = cub.cpRelays # cpRelays only return the first relay on the cub
        old_CTs = cub.GetContents("*.StaCt")
        for old_CT in old_CTs:
            old_CT.Delete()
        old_VTs = cub.GetContents("*.StaVt")
        for old_VT in old_VTs:
            old_VT.Delete()
            
        relay = cub.CreateObject("ElmRelay", "Distance Protection")
        relay.SetAttribute("typ_id", distance_polygonal)
        
        relay_CT = cub.CreateObject("StaCt", "Current Transformer")
        relay_CT.SetAttribute("typ_id", CT)
        relay_CT.SetAttribute("ptapset", 1000) # 1000 primary windings  (1 on secondary by default)
        
        relay_VT = cub.CreateObject("StaVt", "Voltage Transformer")
        relay_VT.SetAttribute("typ_id", VT)
        setRelayVT(relay_VT)
        
        relay.SlotUpdate() # Connects the CT and the VT to the relay
        relay.GetContents("Measurement")[0].Unom = 110
        
        # Rmax = 2*line.R1
        # relay.GetContents("Ph-Ph Polygonal 1")[0].cpRmax = Rmax
        # relay.GetContents("Ph-Ph Polygonal 2")[0].cpRmax = Rmax
        # relay.GetContents("Ph-Ph Polygonal 3")[0].cpRmax = Rmax
        # relay.GetContents("Ph-E Polygonal 1")[0].cpRmax = Rmax
        # relay.GetContents("Ph-E Polygonal 2")[0].cpRmax = Rmax
        # relay.GetContents("Ph-E Polygonal 3")[0].cpRmax = Rmax
        
        # Phi (80 by default)
        if True: # line.loc_name == "Line 01 - 39" or line.loc_name == "Line 09 - 39":
            phi = 85
            relay.GetContents("Ph-Ph Polygonal 1")[0].phi = phi
            relay.GetContents("Ph-Ph Polygonal 2")[0].phi = phi
            relay.GetContents("Ph-Ph Polygonal 3")[0].phi = phi
            relay.GetContents("Ph-E Polygonal 1")[0].phi = phi
            relay.GetContents("Ph-E Polygonal 2")[0].phi = phi
            relay.GetContents("Ph-E Polygonal 3")[0].phi = phi
        
        loadFlow = app.GetFromStudyCase("ComInc")
        loadFlow.Execute() # Performing a loadflow during the setting of a relay can cause issues
        # (PowerFactory doesn't compute the power flow if error in the relay), if it is the case,
        # put the loadFlow before and save the results for each line end to set all the relays
        I = line.GetAttribute("m:I1:bus{}".format(side+1))
        
        # Starting for pilot: start when current increases
        # relay.GetContents("Starting")[0].ip2 = min(2 * I, 1.2 * line.Inom_a)
        relay.GetContents("Starting")[0].ip2 = min(1.5 * I, 1.2 * line.Inom_a)
       
        
        # # Starting, old
        # if IEEE9:
        #     relay.GetContents("Starting")[0].ip2 = 0.2 * line.Inom_a
        #     # relay.GetContents("Starting Backup trip delay")[0].outserv = 1
        # else:
        #     relay.GetContents("Starting")[0].ip2 = 0.6 * line.Inom_a
        #     # relay.GetContents("Starting Backup trip delay")[0].outserv = 0

        # Zone 1
        X1 = 0.85 * line.X1
        R1 = 0.85 * line.R1
        relay.GetContents("Ph-Ph Polygonal 1")[0].cpXmax = X1
        relay.GetContents("Ph-E Polygonal 1")[0].cpXmax = X1
        relay.GetContents("Ph-Ph Polygonal 1")[0].cpRmax = R1
        relay.GetContents("Ph-E Polygonal 1")[0].cpRmax = R1
        relay.GetContents("Ph-Ph Polygonal 1 Delay")[0].Tdelay = 0.02
        relay.GetContents("Ph-E Polygonal 1 Delay")[0].Tdelay = 0.02
        
        relay.GetContents("Pilot Delay")[0].Tdelay = 0.04
        
        # Zone 2
        opposite_cub = line.GetCubicle(1-side) # 1-side = 1 for side = 0, and = 0 for side = 1
        connected_lines = opposite_cub.cBusBar.GetConnectedElements() # Lines connected to the bus were the opposite relay is attached
        for connected_line in connected_lines.copy():
            if connected_line.GetClassName() != 'ElmLne':
                connected_lines.remove(connected_line) # Remove non-line objects
        connected_lines.remove(line) # Remove self
        
        if len(connected_lines) != 0:
            shortest_adj_line = connected_lines[0]
            for connected_line in connected_lines:
                if connected_line.X1 < shortest_adj_line.X1:
                    shortest_adj_line = connected_line
                    
            min_adj_X = shortest_adj_line.X1
            min_adj_R = shortest_adj_line.R1
        else:
            min_adj_X = 0
            min_adj_R = 0
            
        
        X2 = max(0.9*(line.X1 + 0.85*min_adj_X), 1.15*line.X1)
        R2 = max(0.9*(line.R1 + 0.85*min_adj_R), 1.15*line.R1)
        # Z2 = 1.2 * line.X1
        relay.GetContents("Ph-Ph Polygonal 2")[0].cpXmax = X2
        relay.GetContents("Ph-E Polygonal 2")[0].cpXmax = X2
        relay.GetContents("Ph-Ph Polygonal 2")[0].cpRmax = R2
        relay.GetContents("Ph-E Polygonal 2")[0].cpRmax = R2
        relay.GetContents("Ph-Ph Polygonal 2 Delay")[0].Tdelay = 0.2
        relay.GetContents("Ph-E Polygonal 2 Delay")[0].Tdelay = 0.2
        
        # Zone 3
        X3 = (line.X1 + 1.15 * min_adj_X)
        R3 = (line.R1 + 1.15 * min_adj_R)
        # Z3 = 1.8 * line.X1
        relay.GetContents("Ph-Ph Polygonal 3")[0].cpXmax = X3
        relay.GetContents("Ph-E Polygonal 3")[0].cpXmax = X3
        relay.GetContents("Ph-Ph Polygonal 3")[0].cpRmax = R3
        relay.GetContents("Ph-E Polygonal 3")[0].cpRmax = R3
        relay.GetContents("Ph-Ph Polygonal 3 Delay")[0].Tdelay = 0.5
        relay.GetContents("Ph-E Polygonal 3 Delay")[0].Tdelay = 0.5
        
        # Pilot scheme (part 1/1)
        # outputLogic = relay.GetContents("Output Logic")[0].sLogic
        # outputLogic[0] = "PUTT = TRIP" # Enable PUTT scheme, for POTT, use outputLogic[1] = "POTT = TRIP"
        # relay.GetContents("Output Logic")[0].sLogic = outputLogic
        relay.GetContents("Output Logic")[0].sLogic = []
        pilot_scheme = 1
        if pilot_scheme:
            relay.GetContents("Output Logic")[0].aDipset = "30" # PUTT on(high), POTT off
        else:
            relay.GetContents("Output Logic")[0].aDipset = "00" # PUTT off, POTT off
        """
        if pilot_scheme:
            X4 = 1 * line.X1
            R4 = 1 * line.R1 * 2
            relay.GetContents("Ph-Ph Polygonal 4")[0].cpXmax = X4
            relay.GetContents("Ph-E Polygonal 4")[0].cpXmax = X4
            relay.GetContents("Ph-Ph Polygonal 4")[0].cpRmax = R4
            relay.GetContents("Ph-E Polygonal 4")[0].cpRmax = R4
            relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].Tdelay = 0.04
            relay.GetContents("Ph-E Polygonal 4 Delay")[0].Tdelay = 0.04
            
            relay.GetContents("Ph-Ph Polygonal 4")[0].outserv = 0
            relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].outserv = 0        
            relay.GetContents("Ph-E Polygonal 4")[0].outserv = 0
            relay.GetContents("Ph-E Polygonal 4 Delay")[0].outserv = 0
        else:
            relay.GetContents("Ph-Ph Polygonal 4")[0].outserv = 1
            relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].outserv = 1        
            relay.GetContents("Ph-E Polygonal 4")[0].outserv = 1
            relay.GetContents("Ph-E Polygonal 4 Delay")[0].outserv = 1
        """
        
        # Unused zones
        relay.GetContents("Ph-E Polygonal 1")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 1 Delay")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 2")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 2 Delay")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 3")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 3 Delay")[0].outserv = 1
        
        relay.GetContents("Ph-Ph Polygonal 4")[0].outserv = 1
        relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].outserv = 1        
        relay.GetContents("Ph-E Polygonal 4")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 4 Delay")[0].outserv = 1
        
        relay.GetContents("Ph-Ph Polygonal 5")[0].outserv = 1
        relay.GetContents("Ph-Ph Polygonal 5 Delay")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 5")[0].outserv = 1
        relay.GetContents("Ph-E Polygonal 5 Delay")[0].outserv = 1
        
        # relay.outserv = 1

# Pilot scheme (part 2/2)
for line in lines:
    side = 0 # Only have to create a interlink on one side
    cub = line.GetCubicle(side)
    opposite_cub = line.GetCubicle(1-side)
    
    link = cub.CreateObject("ElmRelay", "Interlink")
    link.SetAttribute("typ_id", interlink)
    
    link.SetAttribute("Relay 1", opposite_cub.cpRelays)
    link.SetAttribute("Relay 2", cub.cpRelays)
    
    link.outserv = 1
    

"""
Line overcurrent - NOT USED
"""
# overcurrent = library.GetContents("ProtGeneric")[0].GetContents("F50_F51 Phase overcurrent")[0].GetContents("F50_F51 Phase overcurrent")[0]
# for line in lines:
#     for side in range(2):
#         cub = line.GetCubicle(side)
#         relay = cub.CreateObject("ElmRelay", "Line Overcurrent Protection")
#         relay.SetAttribute("typ_id", overcurrent)
        
#         relay.SlotUpdate() # Connects the CT and the VT to the relay
        
#         relay.GetContents("I>")[0].pcharac = overcurrent.GetContents("Characteristics")[0].GetContents("Definite Time")[0]
        
#         relay.GetContents("I>")[0].Ipset = 1.2 * line.Inom_a
#         relay.GetContents("I>")[0].Tpset = 2
        
#         relay.GetContents("I>>")[0].outserv = 1
#         relay.GetContents("I>>>")[0].outserv = 1
#         relay.GetContents("I>>>>")[0].outserv = 1
        
#         relay.outserv = 1

        
"""
Transformer overcurrent - NOT USED
"""
overcurrent = library.GetContents("ProtGeneric")[0].GetContents("F50_F51 Phase overcurrent")[0].GetContents("F50_F51 Phase overcurrent")[0]
TFOs = app.GetCalcRelevantObjects("*.ElmTr2")
for TFO in TFOs:
    for side in range(2):
        cub = TFO.GetCubicle(side)
        
        old_relay = cub.cpRelays
        while old_relay != None:
            old_relay.Delete()
            old_relay = cub.cpRelays # cpRelays only return the first relay on the cub
        old_CTs = cub.GetContents("*.StaCt")
        for old_CT in old_CTs:
            old_CT.Delete()
            
        # if TFO.loc_name == "Trf 11 - 12" or TFO.loc_name == "Trf 13 - 12":
        #     relay = cub.CreateObject("ElmRelay", "Overcurrent Protection")
        #     relay.SetAttribute("typ_id", overcurrent)
            
        #     relay_CT = cub.CreateObject("StaCt", "Current Transformer")
        #     relay_CT.SetAttribute("typ_id", CT)
        #     relay_CT.SetAttribute("ptapset", 1000) # 1000 primary windings  (1 on secondary by default)
            
        #     relay.SlotUpdate() # Connects the CT and the VT to the relay
            
        #     bus_voltage = cub.GetParent().uknom * 1000 # bus voltage given in kV
        #     Inom = (TFO.Snom_a * 1e6) / bus_voltage / sqrt(3) / 1000 # TFO power in MVA, factor 1000 related to the windings of the CT
            
        #     # relay.GetContents("I>")[0].Ipset = 1.2 * Inom
        #     # relay.GetContents("I>")[0].Tpset = 1
        #     relay.GetContents("I>")[0].outserv = 1
            
        #     relay.GetContents("I>>")[0].Ipset = 1.5 * Inom
        #     relay.GetContents("I>>")[0].Tpset = 1.5 # Caution: should be slower than Z3 for selectivity
            
        #     relay.GetContents("I>>>")[0].Ipset = 5 * Inom
        #     relay.GetContents("I>>>")[0].Tpset = 0.5
            
        #     relay.GetContents("I>>>>")[0].Ipset = 10 * Inom
        #     relay.GetContents("I>>>>")[0].Tpset = 0.2

      
underfrequency = library.GetContents("ProtGeneric")[0].GetContents("F81 Frequency")[0].GetContents("F81 Frequency.TypRelay")[0]
"""
Underfrequency load-shedding
"""

loads = app.GetCalcRelevantObjects("*.ElmLod")
for load in loads:
    cub = load.GetCubicle(0)

    relay = cub.cpRelays
    if relay != None:        
        relay.GetContents("F<1 min V")[0].outserv = 1
        relay.GetContents("F<2 min V")[0].outserv = 1
        relay.GetContents("F<3 min V")[0].outserv = 1
        relay.GetContents("F<4 min V")[0].outserv = 1
        
        relay.GetContents("F<1")[0].Ipsetr = 59
        relay.GetContents("F<1")[0].Tpset = 0.1 # Time dial
        relay.GetContents("F<2")[0].Ipsetr = 58.8
        relay.GetContents("F<2")[0].Tpset = 0.1
        relay.GetContents("F<3")[0].Ipsetr = 58
        relay.GetContents("F<3")[0].Tpset = 0.1
        relay.GetContents("F<4")[0].Ipsetr = 57.5
        relay.GetContents("F<4")[0].Tpset = 0.1
        if load.loc_name == "Load 25":
            relay.GetContents("F<1")[0].Ipsetr = 30 # Cannot be put out of service because stage 1 need initialisation, but setting to trip on 30Hz is equivalent
            relay.GetContents("F<2")[0].Ipsetr = 30

"""
Generator protection
"""

gens = app.GetCalcRelevantObjects("*.ElmSym")
for gen in gens:
    cub = gen.GetCubicle(0)
    
    relay = cub.cpRelays
    if relay != None:
        if relay.loc_name == "Over/Under-Speed Protection":
            relay.GetContents("Meas Freq")[0].Unom = 110
            relay.GetContents("Measurement")[0].Unom = 110
            
            # Underfrequency
            relay.GetContents("F<1")[0].Ipsetr = 57 # Frequency
            relay.GetContents("F<1")[0].Tpset = 0.1 # Time dial
            # Overfrequency
            relay.GetContents("F>1")[0].Ipsetr = 63
            relay.GetContents("F>1")[0].Tpset = 0.1
            #Unused
            relay.GetContents("F>1 min V")[0].outserv = 1
            relay.GetContents("F>2")[0].outserv = 1
            relay.GetContents("F>2 min V")[0].outserv = 1
            relay.GetContents("F>3")[0].outserv = 1
            relay.GetContents("F>3 min V")[0].outserv = 1
            relay.GetContents("F>4")[0].outserv = 1
            relay.GetContents("F>4 min V")[0].outserv = 1
            relay.GetContents("F<1 min V")[0].outserv = 1
            relay.GetContents("F<2")[0].outserv = 1
            relay.GetContents("F<2 min V")[0].outserv = 1
            relay.GetContents("F<3")[0].outserv = 1
            relay.GetContents("F<3 min V")[0].outserv = 1
            relay.GetContents("F<4")[0].outserv = 1
            relay.GetContents("F<4 min V")[0].outserv = 1

# for gen in gens: # Multiply the inertia of all generators by 1/2
#     if gen.loc_name == "G1":
#         continue
#     H = gen.GetAttribute("typ_id").h
#     gen.GetAttribute("typ_id").h = 1/2 * H # Test with 1.5 (except gen 1 for IEEE9)
