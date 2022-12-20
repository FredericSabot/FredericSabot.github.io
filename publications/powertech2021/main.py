import sys
sys.path.append(r"C:\Program Files\DIgSILENT\PowerFactory 2019 SP4\Python\3.7")
import powerfactory as pf

from math import atan2, cos, sin, pi, exp
import math
import random
import time as moduleTime
import warnings
from operator import attrgetter
from datetime import datetime
import csv
import os
from itertools import combinations
from scipy.interpolate import interp1d
import scipy.integrate as integrate
import logging

def sqrt(x):
    if x >= 0:
        return math.sqrt(x)
    if x > -1e-4:
        return 0
    else:
        raise ValueError

def stdDeviation(lst):
    N = len(lst)
    if N != 0:
        deviation = sum([i**2 for i in lst]) / N
        mean = sum(lst) / N
        return sqrt(deviation - mean**2)
    elif N == 1:
        return 999999
    else:
        return -1

snapshotPath = "C:\\Users\\Mon pc\\AppData\\Local\\DIgSILENT\\PowerFactory 2019\\Workspace.fED56xS\\"
start = moduleTime.time()

app = pf.GetApplication()
if app is None:
    raise Exception("getting PowerFactory application failed")
app.Hide()

#app.PrintInfo("Python Script started..")
# app.Show()
user = app.GetCurrentUser()

IEEE9 = 0
if IEEE9:
    project_name = "Nine-bus System"
#     project_name = "Nine-bus System - Distance Protection"
#     project_name = "Nine-bus System-Protections"
else:
    project_name = "39 Bus New England System"

project = user.GetContents(project_name + '.IntPrj')[0]
project.st_retention = 0
project.Activate()

if IEEE9 == 0:
    study_case = app.GetActiveStudyCase()
    if study_case.loc_name != "Power Flow":
        raise Exception("Use the \"Power Flow\" study case to speed up the computations (variables are not saved for graphs in this study case)")

#app.PrintPlain("Hello world!")
# lines = app.GetCalcRelevantObjects("*.ElmLne") # DOUBLE CLIC ITEM => extension is in the title of the pop up
#line = lines[0]
#name = line.loc_name
# value = line.GetAttribute('c:loading') # get value for the loading, DOUBLE CLIC ITEM => hover on a box to get the name of a property (the values can ofc be modified)

#app.PrintPlain("Active Project: " + str(project))

now = datetime.now()
date_time = now.strftime("%Y-%m-%dT%H-%M-%S%z")
results_path = "C:\\Users\\Mon pc\\Desktop\\"
if IEEE9:
    grid = "IEEE9"
else:
    grid = "IEEE39"
dir_name = grid + " - " + date_time


os.mkdir(results_path + "\\" + dir_name)
filename = results_path + "\\" + dir_name + "\\Log.txt"
logging.basicConfig(filename=filename,level=logging.DEBUG,format='%(message)s')
handler = logging.StreamHandler(sys.stdout)
handler.setLevel(logging.DEBUG)
root = logging.getLogger()
root.addHandler(handler)

print = root.info
print("Test")


"""
Dynamic simulation parameters
"""
init = app.GetFromStudyCase("ComInc")
init.iopt_sim = "rms"
init.start = -0.1
init.iopt_show = 0  # Don't display result variables
init.iopt_adapt = 1  # Adaptative time step
init.iopt_adaptreset = 0  # Don't reset adaptative time step on restart
init.iopt_adaptadv = 1 # Advanced step size algorithm
init.itrmxmin = 200 # Default 100

simu = app.GetFromStudyCase("ComSim")

def clearSimEvents():
    faultFolder = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
    cont = faultFolder.GetContents()
    for obj in cont:
        obj.Delete()

def addShcEvent(obj, sec, faultType, position = 50):
    faultFolder = app.GetFromStudyCase("Simulation Events/Fault.IntEvt")
    event = faultFolder.CreateObject("EvtShc", obj.loc_name)
    event.p_target = obj
    event.time = sec
    event.i_shc = faultType  # 0 = 3 phase fault, 4 = faut clearing
    obj.ishclne = 1  # Set available for RMS simulation
    obj.fshcloc = position  #  50 = Fault at the middle of the line

# window = app.GetOutputWindow()
# def getOutput():
#     out = window.GetContent(window.MessageType.Plain) # Plain is events + others (only events considered)
#     window.Clear()
#     return out

"""
Fault definition (3-Ph permanent bolted fault at t = 0)
"""
faultType = 0  # Fault type 0 is 3-Phase short circuit
faulClear = 4  # Fault type 4 is fault clearing
faultTime = 0
clearTime = 0.23 # NOT USED
lines = app.GetCalcRelevantObjects("*.ElmLne")
#shcLine = app.GetCalcRelevantObjects("Line 8-9.ElmLne")[0]

for line in lines:
    relay = line.GetCubicle(0).GetContents("Interlink")[0].outserv = 1 # See comment in def Branch.run()

loadFlow = app.GetFromStudyCase("ComInc")
loadFlow.Execute()

loads = app.GetCalcRelevantObjects("*.ElmLod") # "m:Psum:bus1", "m:Qsum:bus1", "s:xspeed" are interesting param too
initLoad = [i.GetAttribute('m:Psum:bus1') for i in loads]
gens = app.GetCalcRelevantObjects("*.ElmSym")

def checkNMinus1LoadFlow():
    """
    Check that lines are not overloaded in N-1 condition
    NOT USED - The lines thermal capacities are not defined (set to 1kA by default in PowerFactory)
    """
    for line in lines:
        line.outserv = 1
        loadFlow.Execute()

        for otherLine in lines:
            if otherLine == line:
                break
            maxLineI = 0
            for side in range(2):
                I = otherLine.GetAttribute("m:I1:bus{}".format(side+1))
                if I > maxLineI:
                    maxLineI = I
            if maxLineI > 1 * otherLine.Inom_a:
                # print("Overload of line {} during failure of line {},\
                #       current = {}, rating = {}".format(getElementName(otherLine),
                #       getElementName(line), maxLineI, otherLine.Inom_a))
                print(getElementName(otherLine), getElementName(line), maxLineI, otherLine.Inom_a, sep = "\t")
        line.outserv = 0



def evalLoadShed():
    """
    Compute the load shedding as the sum of the load shedding at each bus
    
    The load shedding at one bus is 100% if the voltage of the bus is 0.
    If the voltage is not 0, the load shedding is the sum of the underfrequency
    load shedding actions taken
    """
    loadShedding = []

    for i in range(len(loads)):
        init = initLoad[i]
        end = initLoad[i]

        # Partial underfrequency load shedding
        for event in app.GetFromStudyCase("Simulation Events/Fault.IntEvt").GetContents():
            if event.loc_name[:5] == "Stage": # StageOneLoadShedding or StageTwo* or StageThree*
                if event.p_target == loads[i]:
                    end *= (1+event.dP/100)

        # Complete underfrequency load shedding
        if loads[i].GetAttribute('m:Psum:bus1') == 0:
            end = 0

        loadShedding.append(init - end)

    return sum(loadShedding)

def getProba(event):
    """
    Gives the probability of "event"

    """
    if event == "Timer": return 0.0025
    if event == "Distance Protection": return 0.001
    if event == "Overcurrent Protection": return 0.001
    if event == "Line Overcurrent Protection": return 0.001
    if event == "Underfrequency Load Shedding": return 0.01
    if event == "Over/Under-Speed Protection": return 0  # Failure not considered
    if event == "Interlink": return 0.001

    if event == "Ph-Ph Polygonal 2 Delay": return getProba("Timer")
    if event == "Ph-E Polygonal 2 Delay": return getProba("Timer")
    if event == "Ph-Ph Polygonal 3 Delay": return getProba("Timer")
    if event == "Ph-E Polygonal 3 Delay": return getProba("Timer")

    raise NotImplementedError("Event \"" + str(event) + "\" needs to have its probability defined.")

root_branches = [[]]
root_branches2 = [[]]
current_protection_system_root_branches = root_branches
branchGroups = [] # Branches that were run
branchGroups2 = []
current_protection_system_branchGroups = branchGroups
skipped_branches = [] # Branches who are not expected to contribute to the risk
skipped_branches2 = []
current_protection_system_skipped_branches = skipped_branches

#cutoff = 1e-8
if IEEE9:
    cutoff_risk_fraction = 0.0003
    risk_cutoff = 0.01    
else:
    cutoff_risk_fraction = 0.002 # max test (15h) 0.00005 0.002
    risk_cutoff = 17
    
maximum_risk_bias = 0.05

cutoff_frequency = 5e-6
if cutoff_frequency != 0:
    risk_cutoff = cutoff_frequency * sum(initLoad) / cutoff_risk_fraction
    # cf: proba * sum(initLoad) < cutoff_risk_fraction * reference_risk


skipped_branch_count = 0
run_branch_count = 0
current_parameter_set = 0

class BranchGroup:
    """
    A group of branches that have the same initial event (shcLine) and same failures
    """
    def __init__(self, branch):
        self.branches = [branch]
        self.shcLine = branch.shcLine
        self.end = branch.end
        self.outOfService = branch.outOfService
        # self.proba = branch.proba # Proba of a branch can change with the sampling of parameters (effective end of line length change)
    def addBranch(self, branch):
        if set(branch.outOfService) != set(self.outOfService) or branch.shcLine != self.shcLine or branch.end != self.end:
            raise Exception()
        else:
            self.branches.append(branch)

test_count = 0
class Branch:
    def __init__(self, proba, shcLine, end, position, outOfService, parameter_set,
                 motherBranch=None, protectionSystem=0):
        self.loadShed = 0
        self.shcLine = shcLine
        self.end = end
        self.position = position
        self.parameter_set = parameter_set
        self.events = [] # Contains tuples (time, event) for each parameter_set run
        self.motherBranch = None
        self.daughters = []
        self.ghost = 0
        self.protectionSystem = protectionSystem

        if motherBranch == None:
            self.proba = proba
            self.outOfService = outOfService
        else:
            self.motherBranch = motherBranch
            # motherBranch.daughters.append(self), only added if the branch is run
            self.proba = motherBranch.proba * proba
            self.outOfService = motherBranch.outOfService.copy() + outOfService
        if len(self.outOfService) != len(set(self.outOfService)):
            #               for i in self.outOfService:
            #                    printElement(i)
            #               branches.remove(self)
            raise Exception("Duplicate in self.outOfService")

    @classmethod
    def ghostBranch(cls, trigger, motherBranch, current_branchGroups):
        ghost = cls.__new__(cls)
        ghost.parameter_set = motherBranch.parameter_set
        ghost.shcLine = motherBranch.shcLine
        ghost.end = motherBranch.end
        ghost.position = motherBranch.position
        ghost.motherBranch = motherBranch
        ghost.daughters = []
        ghost.proba = motherBranch.proba * getProba(trigger.loc_name)
        ghost.outOfService = motherBranch.outOfService.copy() + [trigger]
        if len(ghost.outOfService) != len(set(ghost.outOfService)):
            raise Exception("Duplicate in self.outOfService")

        # Ghost branches have the same load shedding as their mother (don't have to be computed)
        # ghost.loadShed = motherBranch.loadShed
        ghost.loadShed = 0 # Saying that ghost branches have the same load shedding as their mother
        # is arbitrary because branches can have =/= mothers depending on how they were created
        # e.g. branch with outofservice = [1,2] could have a mother with outofservice = [1]
        # or a mother with outofservice = [2]
        # However this may change if small proba approx is no longer used !!!
        ghost.events = [] # motherBranch.events
        ghost.ghost = 1
        ghost.protectionSystem = motherBranch.protectionSystem

        ghost.saveResults(current_branchGroups)
        return ghost

    def saveResults(self, current_branchGroups = None):
        if self.motherBranch != None:
            self.motherBranch.daughters.append(self)

        global current_protection_system_branchGroups
        if current_branchGroups == None:
            current_branchGroups = current_protection_system_branchGroups

        duplicate = False
        duplicate_of = None
        for i in current_branchGroups:
            if i.shcLine == self.shcLine and i.end == self.end and set(self.outOfService) == set(i.outOfService): # TODO Include equivalent outOfService
                duplicate = True
                duplicate_of = i
                break
        if duplicate:
            duplicate_of.addBranch(self)
        else:
            current_branchGroups.append(BranchGroup(self))

    def run(self, recursive = 1):
        """
        I could not get the PUTT scheme to work properly, so I set the interlink
        to be in service only for the line with the fault
        """
        interlink = self.shcLine.GetCubicle(0).GetContents("Interlink")[0]
        if interlink not in self.outOfService:
            interlink.outserv = 0
        else:
            interlink.outserv = 1

        self.run2(recursive)
        interlink.outserv = 1

    def run2(self, recursive = 1):
        if self.ghost == 1:
            raise Exception("Ghost branches should not be run")
        clearSimEvents()
        app.ResetCalculation()
        simu.tstop = 10
        addShcEvent(self.shcLine, faultTime, faultType, self.position)
        # addShcEvent(shcLine, clearTime, faulClear, self.position) # Clear the fault if no protection

        if risk_cutoff == 0:
            reference_risk = computeRisk(0) # The risk estimate is computed only from the first run,
                                            # to guarantee that it is non-decreasing
        else:
            reference_risk = risk_cutoff

        if self.proba * sum(initLoad) < cutoff_risk_fraction * reference_risk: # initLoad is the maximum load that could possibly by shed
            current_protection_system_skipped_branches.append(self)
            global skipped_branch_count
            skipped_branch_count += 1
        else:
            global run_branch_count
            run_branch_count += 1
            try:
                local_timer = moduleTime.time()
                for i in self.outOfService:
                    setOutServ(i)
                execute(simu)
                
                equilibrium = 1
                faults = app.GetFromStudyCase("Simulation Events/Fault.IntEvt").GetContents()
                for event in faults:
                    if event.time > 1:
                        equilibrium = 0
                        
                if equilibrium == False:
                    simu.tstop = 60
                    execute(simu)

                self.loadShed = evalLoadShed()
                self.saveResults()

                # Set back to normal
                for i in self.outOfService:
                    setInServ(i)

                test = 0
                faults = app.GetFromStudyCase("Simulation Events/Fault.IntEvt").GetContents()
                for event in faults:
                    breaker = event.p_from
                    if breaker != None:
                        trigger = breaker.GetParent()
                        self.events.append((event.time, trigger))
                        if event.time > 30:
                            test = 1
                if test:
                    global test_count
                    test_count += 1
                    print("Test {}".format(test_count))

                print("Branch computed in {:.2f}s, current time: {:.2f}s, line {} out of {}".format(
                    moduleTime.time()-local_timer, moduleTime.time()-start,
                    lines.index(self.shcLine)+1, len(lines)))

                if moduleTime.time()-local_timer > 60:
                    print(getElementName(self.shcLine))
                    print([getElementName(i) for i in self.outOfService])
                    print(str(self.loadShed/sum(initLoad)*100))

                if recursive:
                    """
                    Run a new simulation assuming that one of the event that occured in this branch
                    failed, for each of the events
                    """
                    for fault in faults:
                        breaker = fault.p_from
                        if breaker != None: # (breaker == None for the line short circuit that we of course do not want to put "out of service")
                            trigger = breaker.GetParent()
                            nonConsideredFailures = ["Over/Under-Speed Protection"] # Generator overspeed not considered (assumed to be perfect as it is of interest for power plant owners)
                            if trigger.cDisplayName not in nonConsideredFailures and trigger not in self.outOfService:

                                duplicate = False
                                for branchGroup in current_protection_system_branchGroups:
                                    if self.shcLine == branchGroup.shcLine and self.end == branchGroup.end and self.parameter_set in [branch.parameter_set for branch in branchGroup.branches]:
                                        if set(self.outOfService + [trigger]) == set(branchGroup.outOfService):
                                            duplicate = True

                                        for failure in branchGroup.outOfService:
                                            try:
                                                """
                                                The next block avoid create branches that have
                                                both a failure and a corresponding "upper" failure
                                                e.g.: a Z3 timer failure and a total relay failure
                                                (which disable the use of the Z3 timer completely)
                                                """
                                                parent = failure.GetParent()
                                            except AttributeError:  # If j has no parent
                                                parent = None
                                            if parent == trigger:
                                                duplicate = True

                                if not duplicate:
                                    global current_parameter_set
                                    Branch(getProba(trigger.loc_name), self.shcLine, self.end, self.position, [trigger], current_parameter_set, self).run()

            except KeyboardInterrupt as e:
                print(e)
                # Set grid back to default state
                for i in self.outOfService:
                    setInServ(i)
                raise Exception("Simulation stopped safely")
            except:
                print("Unknow exception")
                # Set grid back to default state
                for i in self.outOfService:
                    setInServ(i)
                raise Exception("Simulation stopped safely")

delay_and_timers = [[]] # List of the failed Z3 timers and their initial delay (see setOutServ())
def setOutServ(element):
    """
    Set "element" out of service. This function adds additional rules compared to element.outserv = 1
    """
    if element.loc_name[-5:] == "Delay": # Failure of Z3 timer is equivalent to setting the delay to 0s
        timers = [i[1] for i in delay_and_timers[current_parameter_set]]
        if element not in timers:
            delay_and_timers[current_parameter_set].append((element.Tdelay, element))
        element.Tdelay = 0.02
    else: # Other elements are simply put out of service
        element.outserv = 1

def setInServ(element):
    if element.loc_name[-5:] == "Delay": # Failure of Z3 timer is equivalent to setting the delay to 0s
        timers = [i[1] for i in delay_and_timers[current_parameter_set]]
        index = timers.index(element)
        element.Tdelay = delay_and_timers[current_parameter_set][index][0]
    else: # Other elements are simply put out of service
        element.outserv = 0


def computeRisk(parameter_set = None):
    risk = 0
    if parameter_set == None:
        for branchGroup in current_protection_system_branchGroups:
            branchGroup_risk = 0
            for branch in branchGroup.branches:
                branchGroup_risk += branch.proba * branch.loadShed
            risk += branchGroup_risk / len(branchGroups[0].branches)
        return risk

    else:
        for branchGroup in current_protection_system_branchGroups:
            for branch in branchGroup.branches:
                if parameter_set == branch.parameter_set:
                    risk += branch.proba * branch.loadShed
        return risk

def checkRiskThreshold():
    proba = 0
    for i in current_protection_system_skipped_branches:
        proba += i.proba
    if proba * sum(initLoad) > maximum_risk_bias * computeRisk():
        warnings.warn("Consider reducing the cutoff_risk_fraction threshold\
                       \nTotal probability skipped: " + str(proba))

def removeNegligeableGroups():
    """
    Consider only branches that should not have been skipped
    Indeed, because the risk estimate increases during the first MCDET branch, some branches are
    computed even though their contribution to the risk is negligible compared to the final estimate.
    TODO: develop this comment
    """
    global current_protection_system_branchGroups
    global risk_cutoff
    if risk_cutoff == 0:
        risk_cutoff = computeRisk(0)

    for branchGroup in current_protection_system_branchGroups.copy():
        if branchGroup.branches[0].proba * sum(initLoad) < cutoff_risk_fraction * risk_cutoff:
            current_protection_system_branchGroups.remove(branchGroup)

            for branch in branchGroup.branches:
                branch.motherBranch.daughters.remove(branch) # Also removes the negligeable branches
                # from the memory of their mother, to avoid them being rebuilt by completeTree()


def computeStandardError(nbRun):
    if nbRun < 2:
        indicator = 999999**2
    else:
        risk = []
        for parameter_set in range(nbRun):
            parameter_set_risk = 0
            for branchGroup in current_protection_system_branchGroups:
                for branch in branchGroup.branches:
                    if branch.parameter_set == parameter_set:
                        parameter_set_risk += branch.proba * branch.loadShed
            risk.append(parameter_set_risk)

        N = nbRun
        indicator = stdDeviation(risk)**2 / (N-1)

    return sqrt(indicator)

def completeTree():
    removeNegligeableGroups()

    nbProtectionSystems = 2
    if branchGroups2 == []:
        nbProtectionSystems = 1

    for i in range(nbProtectionSystems):
        if i == 0:
            current_branchGroups = branchGroups
        else:
            current_branchGroups = branchGroups2

        for branchGroup in current_branchGroups:
            daughters_triggers = []
            for branch in branchGroup.branches:
                for daughter in branch.daughters:
                    triggers = []
                    for failure in daughter.outOfService:
                        if failure not in branchGroup.outOfService:
                            triggers.append(failure)

                    for trigger in triggers:
                        if trigger not in daughters_triggers:
                            daughters_triggers.append(trigger)

            for branch in branchGroup.branches:
                for trigger in daughters_triggers:
                    triggerIsMissing = True
                    for daughter in branch.daughters:
                        if trigger in daughter.outOfService:
                            triggerIsMissing = False

                    if triggerIsMissing:
                        Branch.ghostBranch(trigger, branch, current_branchGroups)

def execute(simu):
    """ Dummy function for profiling"""
    simu.Execute()

def buildRoots(shcLine, Z3failure = 1, faultLocations=None):
    """
    faultLocations is used to buildRoots with given locations (e.g. to rebuild root
    branches with the same fault positions as the initial set for correlated sampling)

    Z3failure also include interlink failure
    """
    line_fault_frequency_per_km = 0.27/100
    line_fault_frequency = shcLine.dline * line_fault_frequency_per_km

    randomisePositions = 1

    global current_parameter_set
    if current_parameter_set == 0:
        randomisePositions = 0 # So that first run is always completely deterministic

    X = shcLine.X1
    relay_max_X = shcLine.GetCubicle(1).cpRelays.GetContents("Ph-Ph Polygonal 1")[0].cpXmax
    endA = (1-relay_max_X/X) * 100
    if endA < 0.1:
        endA = 0.1
    relay_max_X = shcLine.GetCubicle(0).cpRelays.GetContents("Ph-Ph Polygonal 1")[0].cpXmax
    endB = (relay_max_X/X) * 100
    if endB > 99.9:
        endB = 99.9

    if faultLocations == None:
        if randomisePositions:
            positions = [("A", random.uniform(0.1, endA), 15), # Position cannot be stricly 0, otherwise there is no impedance between the fault and a load, so underfrequency load shedding relay will trip
                     ("M", random.uniform(endA, endB), 70), # Weight assumed constant even though the endA and endB are not
                     ("B", random.uniform(endB, 99.9), 15)] # TODO: consider variable frequencies
        else: # Faults at the middle of the segments
            positions = [("A", 0.1, 15),
                     ("M", 50, 70),
                     ("B", 99.9, 15)]

            # positions = [("A", 0.1, 15),
            #          ("M", 50, 70),
            #          ("B", 99.9, 15)]

        # positions = [("M", 50, 100)]

    else:
        positions = [("A", faultLocations[0], 15),
                     ("M", faultLocations[1], 70),
                     ("B", faultLocations[2], 15)]
        if faultLocations[0] > endA:
            print(faultLocations)
            raise ValueError("Position of fault A")
        if faultLocations[1] < endA or faultLocations[1] > endB:
            print(faultLocations)
            raise ValueError("Position of fault M")
        if faultLocations[2] < endB:
            print(faultLocations)
            raise ValueError("Position of fault B")

    # position = random.uniform(0, 100)

    for i in positions:
        end = i[0]
        position = i[1]
        fraction = i[2]/100
        frequency = line_fault_frequency * fraction

        initial_branch = Branch(frequency, shcLine, end, position, [], current_parameter_set)
        current_protection_system_root_branches[current_parameter_set].append(initial_branch)

        if Z3failure: # also include interlink failure

            # Z2 + Z3 timer failure + interlink failure + combinations of them
            clearSimEvents()
            app.ResetCalculation()
            simu.tstop = 0.025
            addShcEvent(shcLine, faultTime, faultType, position)
            execute(simu)

            vulnerable_relays = []

            # Check which relays will incorectly trip if their Z2 or Z3 timer is bypassed
            for line in lines:
                if line == shcLine:
                    if end != "M": # == "A" or "B"
                        # Interlink failure (only important for fault at the end of the line so not considered for the middle)
                        cub = line.GetCubicle(0) # Interlink is only on side 0
                        vulnerable_relays.append((cub.GetContents("Interlink")[0], 1))

                    continue # No Z3 failure to check on the line itself

                for side in range(2):
                    I = line.GetAttribute("m:I1:bus{}".format(side+1))
                    V = line.GetAttribute("n:u1:bus{}".format(side+1)) * line.GetCubicle(0).cBusBar.uknom
                    P = line.GetAttribute("m:Psum:bus{}".format(side+1))
                    Q = line.GetAttribute("m:Qsum:bus{}".format(side+1))
                    S = sqrt(P**2+Q**2)
                    Z = V**2/S
                    theta = atan2(Q, P)
                    X = Z * sin(theta)
                    R = Z * cos(theta)

                    cub = line.GetCubicle(side)
                    Z2_trigger = False
                    Z3_trigger = False
                    parameters = cub.cpRelays.GetContents()
                    e_directional = cub.cpRelays.GetContents("Ground Directional")[0].phi
                    ph_directional = cub.cpRelays.GetContents("Phase Directional")[0].phi
                    for param in parameters:
                        if param.outserv == 1:
                            continue
                        if param.loc_name[:3] == "Ph-" and (param.loc_name[-1] == "2" or param.loc_name[-1] == "3"):
                            if I < cub.cpRelays.GetContents("Starting")[0].ip2: # Current below the pick up current
                                continue
                            X_relay = param.cpXmax
                            R_relay = param.cpRmax
                            relay_angle = param.phi

                            condition1 = (X < X_relay)
                            if param.loc_name[3] == "P": # Ph-Ph
                                condition2 = (0-ph_directional*pi/180 < theta and theta < pi-ph_directional*pi/180)
                            else: # Ph-E
                                condition2 = (0-e_directional*pi/180 < theta and theta < pi-e_directional*pi/180)
                            max_R_relay = R_relay + X*cos(relay_angle/180*pi)
                            condition3 = (R < max_R_relay)
                            min_R_relay = -R_relay + X*cos(relay_angle/180*pi)
                            condition4 = (R > min_R_relay)

                            if condition1 and condition2 and condition3 and condition4:
                                # Apparent impedance is in the zone 2 or 3 of the relay
                                if param.loc_name[-1] == "2": # In zone 2
                                    Z2_trigger = True
                                else: # In zone 3
                                    Z3_trigger = True
                    if Z2_trigger:
                        vulnerable_relays.append((cub.cpRelays, 2))
                    elif Z3_trigger:
                        vulnerable_relays.append((cub.cpRelays, 3))

            # https://docs.python.org/3/library/itertools.html#itertools.combinations
            # All combinations of the Z2 or Z3 timer failures
            for i in range(1, len(vulnerable_relays) + 1):
                for combination in combinations(vulnerable_relays, i):
                    failures = []
                    proba = 1
                    for vulnerable_relay in combination:
                        relay = vulnerable_relay[0]
                        zone = vulnerable_relay[1]
                        proba *= getProba("Timer")
                        if zone == 2:
                            proba *= 2 # If the apparent impedance is in the zone 2, it is also in the zone 3
                            # so the failure of the Z2 timer or the Z3 timer are equivalent
                            # We thus only compute the case of the Z2 timer and multiply its probability by 2
                            failures.append(relay.GetContents("Ph-Ph Polygonal 2 Delay")[0])
                            # failures.append(relay.GetContents("Ph-E Polygonal 2 Delay")[0])
                            # TODO: put back on if consider Ph-E faults, but don't forget that completeTree is bugged with this
                        elif zone == 3:
                            failures.append(relay.GetContents("Ph-Ph Polygonal 3 Delay")[0])
                            # failures.append(relay.GetContents("Ph-E Polygonal 3 Delay")[0])
                        else:
                            failures.append(relay)
                    current_protection_system_root_branches[current_parameter_set].append(Branch(proba, shcLine, end, position, failures, current_parameter_set, initial_branch))

            if risk_cutoff == 0:
                current_protection_system_root_branches[current_parameter_set].sort(key=attrgetter("proba"), reverse=True)


def runAll(recursive = 1, Z3failure = 1):
    for line in lines:
        buildRoots(line, Z3failure)
    numberDirectSimu = len(current_protection_system_root_branches[current_parameter_set])
    runAllTimer = moduleTime.time()
    for i in current_protection_system_root_branches[current_parameter_set]:
        if i.loadShed != 0:
            print("Warning: branch already run")
            break
        i.run(recursive)
    print("\nRun time: " + str(moduleTime.time() - runAllTimer))
    global skipped_branch_count
    global run_branch_count
    print("Number of \"direct\" simulations: " + str(numberDirectSimu))
    print("Total simulation run: " + str(run_branch_count))
    print("Number of skipped branches: " + str(skipped_branch_count))
    print("Mean total risk: " + str(computeRisk()))
    print("Standard error of risk: " + str(computeStandardError(current_parameter_set)))
    try:
        print("Standard error in %: {:.2f}".format(computeStandardError(current_parameter_set)/computeRisk()*100))
    except ZeroDivisionError:
        pass
    skipped_branch_count = 0
    run_branch_count = 0
    print("")

save_Ph = [[]]
save_underfrequency = [[]]
save_overspeed = [[]]

def filterRelevantParameters(parameters):
    """
    Sort parameters (to simplify saving/loading parameters) and filter unused ones (outserv == 1)
    """
    for param in sorted(parameters, key=attrgetter("loc_name")):
        if param.outserv == 0:
            yield param

def saveInitialParameters():
    if save_Ph != [[]] or save_underfrequency != [[]] or save_overspeed != [[]]:
        raise Exception("saveInitialParameters() can only be called once")

    # Distance protection parameters
    for i in range(len(lines)):
        line_save_Ph = []
        save_Ph[0].append(line_save_Ph)
        for side in range(2):
            cub_save_Ph = []
            line_save_Ph.append(cub_save_Ph)
            cub = lines[i].GetCubicle(side)
            parameters = cub.cpRelays.GetContents()
            for param in filterRelevantParameters(parameters):
                if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] != "Delay": # True for Ph-Ph Polygonal 1 to 3 and Ph-E Polygonal 1 to 5 and not for "Ph-E Polygonal 1 Delay"
                    cub_save_Ph.append(param.Rmax)
                    cub_save_Ph.append(param.Xmax)
                if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] == "Delay":
                    cub_save_Ph.append(param.Tdelay)

    # Underfrequency load shedding parameters
    underfrequency = []
    for load in loads:
        if load.GetCubicle(0).cpRelays != None:
            underfrequency.append(load.GetCubicle(0).cpRelays)
    for relay in underfrequency:
        relay_save = []
        save_underfrequency[0].append(relay_save)
        parameters = relay.GetContents()
        for param in filterRelevantParameters(parameters):
            if param.loc_name[0] == "F":
                relay_save.append(param.Ipsetr)
                relay_save.append(param.Tpset)

    # Over/under protection of generators parameters
    overspeed = []
    for gen in gens:
        if gen.GetCubicle(0).cpRelays != None:
            overspeed.append(gen.GetCubicle(0).cpRelays)
    for relay in overspeed:
        relay_save = []
        save_overspeed[0].append(relay_save)
        parameters = relay.GetContents()
        for param in filterRelevantParameters(parameters):
            if param.loc_name[0] == "F":
                relay_save.append(param.Ipsetr)
                relay_save.append(param.Tpset)


def incrementCurrentParameterSet():
    global current_parameter_set
    current_parameter_set += 1
    save_Ph.append([])
    save_underfrequency.append([])
    save_overspeed.append([])

    delay_and_timers.append([])

gauss = 1 # gauss = 1 is noise only, gauss = 0 is optimisation (without noise)
distance_variance = 0.01
distance_optimisation_range = 0.1
frequency_variance = 0.001

delay_flat_variance = 0.5/60 # Flat delay as opposed to the other parameters which are
                             # in percentage. The uncertainty is taken as half a cycle (at 60 Hz)

if gauss == 0:
    distance_variance = 0
    frequency_variance = 0
    delay_flat_variance = 0

def randomiseParameters():
    if save_Ph == [[]] or save_underfrequency == [[]]: #  or save_overspeed == [[]]
        raise Exception("Please call saveInitialParameters() before randomising parameters")

    incrementCurrentParameterSet()
    # Distance protection parameters
    for line in range(len(lines)):
        line_save_Ph = []
        save_Ph[current_parameter_set].append(line_save_Ph)
        for side in range(2):
            cub_save_Ph = []
            line_save_Ph.append(cub_save_Ph)
            cub = lines[line].GetCubicle(side)
            parameters = cub.cpRelays.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] != "Delay":
                    # Resistance treshold
                    if gauss or (param.loc_name != "Ph-Ph Polygonal 3" \
                                  and param.loc_name != "Ph-E Polygonal 3"):
                        cub_save_Ph.append(
                            save_Ph[0][line][side][j] * (1 + random.gauss(0, distance_variance)))
                    else:
                        cub_save_Ph.append(
                            save_Ph[0][line][side][j] * (1 + random.gauss(0, distance_optimisation_range)))
                    test = param.Rmax # check that the param exists
                    param.Rmax = save_Ph[current_parameter_set][line][side][j]
                    j += 1
                    # Impedance treshold
                    if gauss or (param.loc_name != "Ph-Ph Polygonal 3" \
                                  and param.loc_name != "Ph-E Polygonal 3"): # We do not try to optimise the
                        # Z1 range as ~85% is always the best
                        cub_save_Ph.append(
                            save_Ph[0][line][side][j] * (1 + random.gauss(0, distance_variance)))
                    else:
                        cub_save_Ph.append(
                            save_Ph[0][line][side][j] * (1 + random.gauss(0, distance_optimisation_range)))
                    test = param.Xmax
                    param.Xmax = save_Ph[current_parameter_set][line][side][j]
                    j += 1
                elif param.loc_name[:3] == "Ph-" and param.loc_name[-5:] == "Delay":
                    t = save_Ph[0][line][side][j] + random.gauss(0, delay_flat_variance)
                    cub_save_Ph.append(max(0.01, t))
                    test = param.Tdelay
                    param.Tdelay = save_Ph[current_parameter_set][line][side][j]
                    j += 1
    # Underfrequency load shedding parameters
    for load in range(len(loads)):
        underf_save = []
        save_underfrequency[current_parameter_set].append(underf_save)
        relay = loads[load].GetCubicle(0).cpRelays
        if relay != None:
            parameters = relay.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[0] == "F":
                    underf_save.append(
                        save_underfrequency[0][load][j] * (1 + random.gauss(0, frequency_variance)))
                    test = param.Ipsetr
                    param.Ipsetr = save_underfrequency[current_parameter_set][load][j]
                    j += 1
                    t = save_underfrequency[0][load][j] + random.gauss(0, delay_flat_variance)
                    underf_save.append(max(0.01, t))
                    test = param.Tpset
                    param.Tpset = save_underfrequency[current_parameter_set][load][j]
                    j += 1

    # Over/under protection of generators parameters
    for gen in range(len(gens)):
        overspeed_save = []
        save_overspeed[current_parameter_set].append(overspeed_save)
        relay = gens[gen].GetCubicle(0).cpRelays
        if relay != None:
            parameters = relay.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[0] == "F":
                    overspeed_save.append(
                        save_overspeed[0][gen][j] * (1 + random.gauss(0, frequency_variance)))
                    test = param.Ipsetr
                    param.Ipsetr = save_overspeed[current_parameter_set][gen][j]
                    j += 1
                    t = save_overspeed[0][gen][j] + random.gauss(0, delay_flat_variance)
                    overspeed_save.append(max(0.01, t))
                    test = param.Tpset
                    param.Tpset = save_overspeed[current_parameter_set][gen][j]
                    j += 1


def setDefaultParameters():
    """
    Set parameters back to initial state
    """
    setParameters(0)

def setParameters(parameter_set):
    """
    Set parameters to the values of the parameter set parameter_set.
    Can be used for manual debugging/verification purposes.
    """
    # Distance protection parameters
    for line in range(len(lines)):
        for side in range(2):
            cub = lines[line].GetCubicle(side)
            parameters = cub.cpRelays.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] != "Delay":
                    param.Rmax = save_Ph[parameter_set][line][side][j]
                    j += 1
                    param.Xmax = save_Ph[parameter_set][line][side][j]
                    j += 1
                if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] == "Delay":
                    param.Tdelay = save_Ph[parameter_set][line][side][j]
                    j += 1

    # Underfrequency load shedding parameters
    for load in range(len(loads)):
        relay = loads[load].GetCubicle(0).cpRelays
        if relay != None:
            parameters = relay.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[0] == "F":
                    param.Ipsetr = save_underfrequency[parameter_set][load][j]
                    j += 1
                    param.Tpset = save_underfrequency[parameter_set][load][j]
                    j += 1

    # Over/under protection of generators parameters
    for gen in range(len(gens)):
        relay = gens[gen].GetCubicle(0).cpRelays
        if relay != None:
            parameters = relay.GetContents()
            j = 0
            for param in filterRelevantParameters(parameters):
                if param.loc_name[0] == "F":
                    param.Ipsetr = save_overspeed[parameter_set][gen][j]
                    j += 1
                    param.Tpset = save_overspeed[parameter_set][gen][j]
                    j += 1


def parameterSetsToCSV():
    if save_Ph == [[]]:
        saveInitialParameters()
    filename = results_path + "\\" + dir_name + "\\Parameter sets.csv"
    nbRun = len(current_protection_system_branchGroups[0].branches)

    with open(filename, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile, dialect='excel', delimiter=';')
        # Distance protection parameters
        writer.writerow(["Element"] + [str(param) for param in range(len(current_protection_system_branchGroups[0].branches))])
        for line in range(len(lines)):
            for side in range(2):
                cub = lines[line].GetCubicle(side)
                parameters = cub.cpRelays.GetContents()
                j = 0
                for param in filterRelevantParameters(parameters):
                    if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] != "Delay":
                        writer.writerow([getElementName(param) + "Rmax"] +
                                         [save_Ph[parameter_set][line][side][j] for parameter_set in range(nbRun)])
                        j += 1
                        writer.writerow([getElementName(param) + "Xmax"] +
                                         [save_Ph[parameter_set][line][side][j] for parameter_set in range(nbRun)])
                        j += 1
                    if param.loc_name[:3] == "Ph-" and param.loc_name[-5:] == "Delay":
                        writer.writerow([getElementName(param) + "Tdelay"] +
                                         [save_Ph[parameter_set][line][side][j] for parameter_set in range(nbRun)])
                        j += 1

        # Underfrequency load shedding parameters
        for load in range(len(loads)):
            relay = loads[load].GetCubicle(0).cpRelays
            if relay != None:
                parameters = relay.GetContents()
                j = 0
                for param in filterRelevantParameters(parameters):
                    if param.loc_name[0] == "F":
                        writer.writerow([getElementName(param) + "Ipsetr"] +
                                         [save_underfrequency[parameter_set][load][j] for parameter_set in range(nbRun)])
                        j += 1
                        writer.writerow([getElementName(param) + "Tpset"] +
                                         [save_underfrequency[parameter_set][load][j] for parameter_set in range(nbRun)])
                        j += 1

        # Over/under protection of generators parameters
        for gen in range(len(gens)):
            relay = gens[gen].GetCubicle(0).cpRelays
            if relay != None:
                parameters = relay.GetContents()
                j = 0
                for param in filterRelevantParameters(parameters):
                    if param.loc_name[0] == "F":
                        writer.writerow([getElementName(param) + "Ipsetr"] +
                                         [save_overspeed[parameter_set][gen][j] for parameter_set in range(nbRun)])
                        j += 1
                        writer.writerow([getElementName(param) + "Tpset"] +
                                         [save_underfrequency[parameter_set][load][j] for parameter_set in range(nbRun)])
                        j += 1

def runMCDET(nbRuns, recursive = 1, Z3failure = 1):
    runAll(recursive, Z3failure)  # First run with default parameters
    removeNegligeableGroups()
    saveInitialParameters()
    try:
        for i in range(nbRuns-1):  # -1 because one run with default parameter already done
            randomiseParameters()
            current_protection_system_root_branches.append([])
            print("\nRunning set of branches no: " + str(current_parameter_set))
            runAll(recursive, Z3failure)
        setDefaultParameters()
    except Exception as e:
        print(e)
        setDefaultParameters()
        raise Exception("Simulation stopped safely")
    except:
        print("Unknow exception")
        setDefaultParameters()
        raise Exception("Simulation stopped safely")


def printLoadShed():
    for branchGroup in current_protection_system_branchGroups:
        print("Branch: " + branchGroup.shcLine.loc_name)
        for failure in branchGroup.outOfService:
            print(getElementName(failure))
        for branch in branchGroup.branches:
            print("{:.2f}".format(branch.loadShed/sum(initLoad) * 100))
        print("")

def getElementName(element):
    """
    Returns only the part of element's name that is after the Grid.ElmNet part
    """
    out = ''
    full_name = element.GetFullName().split("\\")
    for i in reversed(full_name):
        if i[-len('.ElmNet'):] == '.ElmNet':
            break
        out = i.split(".")[0] + " " + out # Remove the extention name then add to the final string
    return out

def computeImportanceMeasures(toCSV = 0):
    completeTree()
    nb_run = len(branchGroups[0].branches)

    elements = [] # List of elements that have failed in at least one simulation
    for branchGroup in current_protection_system_branchGroups:
        for failure in branchGroup.outOfService:
            if failure not in elements:
                elements.append(failure)
    RRW = []

    for element in elements:
        if element.loc_name[-22:] == "Ph-E Polygonal 3 Delay" or \
            element.loc_name[-22:] == "Ph-E Polygonal 2 Delay": # Duplicate of Ph-Ph Polygonal {} Delay
            continue

        risk_reduction_by_parameter = [0] * nb_run
        cost_reduction_by_parameter = [0] * nb_run
        for branchGroup in current_protection_system_branchGroups:
            if element in branchGroup.outOfService:
                for branch in branchGroup.branches:
                    risk_reduction_by_parameter[branch.parameter_set] += branch.proba * branch.loadShed
                    cost_reduction_by_parameter[branch.parameter_set] += branch.proba * loadShedToCost(branch.loadShed)

        risk_reduction = sum(risk_reduction_by_parameter) / nb_run
        cost_reduction = sum(cost_reduction_by_parameter) / nb_run

        RRW.append([element, risk_reduction, stdDeviation(risk_reduction_by_parameter)] +
                   risk_reduction_by_parameter +
                   [cost_reduction, stdDeviation(cost_reduction_by_parameter)] +
                   cost_reduction_by_parameter)

    if toCSV:
        filename = results_path + "\\" + dir_name + "\\Importance measures" + ".csv"

        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', delimiter=';')
            writer.writerow(["Element", "Mean RRW (MW/y)", "RRW deviation (MW/y)"] +
                            ["RRW {} (MW/y)".format(param+1) for param in range(nb_run)] +
                            ["Mean RRW ($/y)", "RRW deviation ($/y)"] +
                            ["RRW {} ($/y)".format(param+1) for param in range(nb_run)])
            for i in RRW:
                writer.writerow([getElementName(i[0])] + i[1:])
    return RRW

def computeLineImportance(toCSV = 0):
    completeTree()
    nb_run = len(branchGroups[0].branches)

    lineImportances = []
    for line in lines:
        risk_by_parameter = [0] * nb_run
        cost_by_parameter = [0] * nb_run
        for branchGroup in current_protection_system_branchGroups:
            if branchGroup.shcLine == line:
                for branch in branchGroup.branches:
                    risk_by_parameter[branch.parameter_set] += branch.proba * branch.loadShed
                    cost_by_parameter[branch.parameter_set] += branch.proba * loadShedToCost(branch.loadShed)
        risk = sum(risk_by_parameter) / nb_run
        cost = sum(cost_by_parameter) / nb_run

        lineImportances.append([line, risk, stdDeviation(risk_by_parameter)] +
                               risk_by_parameter + [cost, stdDeviation(cost_by_parameter)] +
                               cost_by_parameter)
    if toCSV:
        filename = results_path + "\\" + dir_name + "\\Line importance" + ".csv"

        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', delimiter=';')
            writer.writerow(["Line", "Mean risk", "Risk deviation"] +
                            ["Risk {}".format(i+1) for i in range(nb_run)] +
                            ["Cost", "Cost deviation"] + ["Cost {}".format(i+1) for i in range(nb_run)])
            for i in lineImportances:
                writer.writerow([getElementName(i[0])] + i[1:])
    return lineImportances

def computeEndImportance(toCSV = 0):
    completeTree()
    nb_run = len(branchGroups[0].branches)

    endImportances = []
    for end in ["A", "M", "B"]:
        risk_by_parameter = [0] * nb_run
        cost_by_parameter = [0] * nb_run
        for branchGroup in current_protection_system_branchGroups:
            if branchGroup.end == end:
                for branch in branchGroup.branches:
                    risk_by_parameter[branch.parameter_set] += branch.proba * branch.loadShed
                    cost_by_parameter[branch.parameter_set] += branch.proba * loadShedToCost(branch.loadShed)
        risk = sum(risk_by_parameter) / nb_run
        cost = sum(cost_by_parameter) / nb_run

        endImportances.append([end, risk, stdDeviation(risk_by_parameter)] +
                               risk_by_parameter + [cost, stdDeviation(cost_by_parameter)] +
                               cost_by_parameter)
    if toCSV:
        filename = results_path + "\\" + dir_name + "\\End importance" + ".csv"

        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', delimiter=';')
            writer.writerow(["End", "Mean risk", "Risk deviation"] +
                            ["Risk {}".format(i+1) for i in range(nb_run)] +
                            ["Cost", "Cost deviation"] + ["Cost {}".format(i+1) for i in range(nb_run)])
            for i in endImportances:
                writer.writerow(i)
    return endImportances



def loadShedToCost(loadShed):
    """
    Compute Value of Lost Load (VoLL)
    Derived from Pierre Henneaux and  Daniel S. Kirschen, "Probabilistic Security
    Analysis of Optimal Transmission Switching"

    """
    time = [1/60, 20/60, 1, 4, 8, 24]
    cost = [571, 61, 39, 30, 27, 13]
    for i in range(len(cost)):
        cost[i] *= 1000 # from kWh to MWh

    def interpolVoLL(t):
        if t >= time[0]:
            return interp1d(time, cost)(t)
        else:
            return interp1d(time, cost)(time[0])

    # interpol = interp1d(time, cost)

    LOL = loadShed/sum(initLoad) * 100
    H = 0.1419 * LOL + 0.6482

    k = 3
    return integrate.quad(lambda t: k/H*exp(-k*t/H) * loadShed * t * interpolVoLL(t), 0, H/2)[0]


def exportToCSV():
    completeTree()
    filename = results_path + "\\" + dir_name + "\\Main contributors" + ".csv"

    nb_run = len(branchGroups[0].branches)

    nbProtectionSystems = 2
    if branchGroups2 == []:
        nbProtectionSystems = 1

    for i in range(nbProtectionSystems):
        if nbProtectionSystems == 1:
            filename = results_path + "\\" + dir_name + "\\Main contributors" + ".csv"
        else:
            filename = results_path + "\\" + dir_name + "\\Main contributors {}".format(i+1) + ".csv"

        with open(filename, 'w', newline='') as csvfile:
            if i == 0:
                current_branchGroups = branchGroups
            else:
                current_branchGroups = branchGroups2
            writer = csv.writer(csvfile, dialect='excel', delimiter=';')
            writer.writerow(["Mean risk (Mw/y)", "Risk contribution (%)", "Risk deviation",
                             "Mean cost ($/y)", "Cost deviation", "Mean frequency", "Line", "End"] +
                             ["Fault position", "Local risk", "Local risk contribution(%)", "Local cost",
                             "Load shed (%)", "Parameter set", "Ghost"] * nb_run +
                            ["Out of service"]*3 +
                            ["Failure category {}".format(i+1) for i in range(3)] + ["Merged category"])
            for branchGroup in current_branchGroups:
                branchGroup.branches.sort(key=attrgetter("parameter_set"))

                branchGroup_risks = [branch.proba * branch.loadShed for branch in branchGroup.branches]
                branchGroup_risk = sum(branchGroup_risks)/nb_run
                branchGroup_deviation = stdDeviation(branchGroup_risks)

                total_risk = computeRisk()

                branchGroup_costs = [branch.proba * loadShedToCost(branch.loadShed) for branch in branchGroup.branches]
                branchGroup_cost = sum(branchGroup_costs) / nb_run
                branchGroup_cost_deviation= stdDeviation(branchGroup_costs)

                local_stuff = []
                for branch in branchGroup.branches:
                    total_risk = computeRisk(branch.parameter_set)
                    if total_risk == 0:
                        total_risk = 1e-10 # Avoid ZeroDivisionError
                    local_stuff.append(branch.position)
                    local_stuff.append(branch.proba*branch.loadShed)
                    local_stuff.append(branch.proba*branch.loadShed/total_risk*100)
                    local_stuff.append(branch.proba*loadShedToCost(branch.loadShed))
                    local_stuff.append(branch.loadShed/sum(initLoad)*100)
                    local_stuff.append(branch.parameter_set)
                    local_stuff.append(branch.ghost)

                failures = []
                for failure in branchGroup.outOfService:
                    if failure.loc_name[-22:] != "Ph-E Polygonal 3 Delay" and failure.loc_name[-22:] != "Ph-E Polygonal 2 Delay": # Duplicate of Ph-Ph Polygonal * Delay
                        failures.append(getElementName(failure))
                while len(failures) < 3:
                    failures.append("")

                failure_categories = []
                for failure in failures:
                    if "Distance Protection" in failure:
                        if "Delay" in failure:
                            failure_categories.append("Delay")
                        else:
                            failure_categories.append("Distance Protection")
                    elif "Underfrequency Load Shedding" in failure:
                        failure_categories.append("Underfrequency Load Shedding")
                    elif failure == "":
                        failure_categories.append("")
                    elif "Interlink" in failure:
                        failure_categories.append("Interlink")
                    else:
                        failure_categories.append("Other")
                merged_category = "-".join([i for i in sorted(failure_categories) if i]) # if i removes the empty elements ("")

                branchGroup_events = []
                for event in branchGroup.branches[0].events:
                    if getElementName(event[1]) not in branchGroup_events:
                        branchGroup_events.append(event[0])
                        branchGroup_events.append(getElementName(event[1]))

                writer.writerow([branchGroup_risk, branchGroup_risk/total_risk * 100, branchGroup_deviation] +
                                [branchGroup_cost, branchGroup_cost_deviation] +
                                [sum(branch.proba for branch in branchGroup.branches)/len(branchGroup.branches)] +
                                [getElementName(branchGroup.shcLine), branchGroup.end] +
                                local_stuff + failures +
                                failure_categories + [merged_category] + branchGroup_events)

        # Printing of events
        for param in range(current_parameter_set+1):
            if nbProtectionSystems == 1:
                filename = results_path + "\\" + dir_name + "\\Events {}".format(param) + ".csv"
            else:
                filename = results_path + "\\" + dir_name + "\\Events {} - {}".format(i+1, param) + ".csv"

            with open(filename, 'w', newline='') as csvfile:
                if i == 0:
                    current_branchGroups = branchGroups
                else:
                    current_branchGroups = branchGroups2
                writer = csv.writer(csvfile, dialect='excel', delimiter=';')
                writer.writerow(["Line", "End"] + ["Out of service"]*3 +
                            ["Failure category {}".format(i+1) for i in range(3)] + ["Merged category"])
                for branchGroup in current_branchGroups:
                    branchGroup.branches.sort(key=attrgetter("parameter_set"))

                    branchGroup_events = []
                    for event in branchGroup.branches[param].events:
                        if getElementName(event[1]) not in branchGroup_events:
                            branchGroup_events.append(event[0])
                            branchGroup_events.append(getElementName(event[1]))

                    failures = []
                    for failure in branchGroup.outOfService:
                        if failure.loc_name[-22:] != "Ph-E Polygonal 3 Delay" and failure.loc_name[-22:] != "Ph-E Polygonal 2 Delay": # Duplicate of Ph-Ph Polygonal * Delay
                            failures.append(getElementName(failure))
                    while len(failures) < 3:
                        failures.append("")

                    failure_categories = []
                    for failure in failures:
                        if "Distance Protection" in failure:
                            if "Delay" in failure:
                                failure_categories.append("Delay")
                            else:
                                failure_categories.append("Distance Protection")
                        elif "Underfrequency Load Shedding" in failure:
                            failure_categories.append("Underfrequency Load Shedding")
                        elif failure == "":
                            failure_categories.append("")
                        elif "Interlink" in failure:
                            failure_categories.append("Interlink")
                        else:
                            failure_categories.append("Other")
                    merged_category = "-".join([i for i in sorted(failure_categories) if i]) # if i removes the empty elements ("")

                    writer.writerow([getElementName(branchGroup.shcLine), branchGroup.end] +
                                failures + failure_categories + [merged_category] + branchGroup_events)


    # Comparison between the two protection systems
    if nbProtectionSystems == 2:
        filename = results_path + "\\" + dir_name + "\\Main contributors - Comparison" + ".csv"
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', delimiter=';')
            writer.writerow(["Frequency", "Line", "End",
                 "Risk 1", "Deviation 1", "Cost 1", "Cost Deviation 1", "Mean load shed 1 (%)",
                 "Risk 2", "Deviation 2", "Cost 2", "Cost Deviation 2", "Mean load shed 2 (%)",
                 "Delta risk", "Delta risk deviation"] +
                ["Delta risk {}".format(i+1) for i in range(nb_run)] +
                ["Delta cost", "Delta cost deviation"] +
                ["Delta cost {}".format(i+1) for i in range(nb_run)] +
                ["Out of service"] * 3 +
                ["Failure category {}".format(i+1) for i in range(nb_run)] + ["Merged category"])

            for branchGroup in branchGroups:
                compareTo = None
                for branchGroup2 in branchGroups2:
                    if set(branchGroup2.outOfService) == set(branchGroup.outOfService) and \
                    branchGroup2.shcLine == branchGroup.shcLine and branchGroup2.end == branchGroup.end:
                        compareTo = branchGroup2

                branchGroup.branches.sort(key=attrgetter("parameter_set"))

                branchGroup_risks = [branch.proba * branch.loadShed for branch in branchGroup.branches]
                branchGroup_risk = sum(branchGroup_risks)/nb_run
                branchGroup_deviation = stdDeviation(branchGroup_risks)

                branchGroup_costs = [branch.proba * loadShedToCost(branch.loadShed) for branch in branchGroup.branches]
                branchGroup_cost = sum(branchGroup_costs) / nb_run
                branchGroup_cost_deviation= stdDeviation(branchGroup_costs)


                branchGroup2_risk = 0
                branchGroup2_deviation = 0
                branchGroup2_cost = 0
                branchGroup2_cost_deviation = 0
                if compareTo != None:
                    compareTo.branches.sort(key=attrgetter("parameter_set"))

                    branchGroup2_risks = [branch.proba * branch.loadShed for branch in compareTo.branches]
                    branchGroup2_risk = sum(branchGroup2_risks)/nb_run
                    branchGroup2_deviation = stdDeviation(branchGroup2_risks)

                    branchGroup2_costs = [branch.proba * loadShedToCost(branch.loadShed) for branch in compareTo.branches]
                    branchGroup2_cost = sum(branchGroup2_costs) / nb_run
                    branchGroup2_cost_deviation= stdDeviation(branchGroup2_costs)


                delta_risks = [0] * nb_run
                for branch in branchGroup.branches:
                        delta_risks[branch.parameter_set] += branch.proba * branch.loadShed
                if compareTo != None:
                    for branch in compareTo.branches:
                        delta_risks[branch.parameter_set] -= branch.proba * branch.loadShed # Minus sign

                delta_risk_deviation = stdDeviation(delta_risks)


                delta_costs = [0] * nb_run
                for branch in branchGroup.branches:
                        delta_costs[branch.parameter_set] += branch.proba * loadShedToCost(branch.loadShed)
                if compareTo != None:
                    for branch in compareTo.branches:
                        delta_costs[branch.parameter_set] -= branch.proba * loadShedToCost(branch.loadShed)

                delta_cost_deviation = stdDeviation(delta_costs)


                failures = []
                for failure in branchGroup.outOfService:
                    if failure.loc_name[-22:] != "Ph-E Polygonal 3 Delay" and failure.loc_name[-22:] != "Ph-E Polygonal 2 Delay": # Duplicate of Ph-Ph Polygonal * Delay
                        failures.append(getElementName(failure))
                while len(failures) < 3:
                    failures.append("")

                failure_categories = []
                for failure in failures:
                    if "Distance Protection" in failure:
                        if "Delay" in failure:
                            failure_categories.append("Delay")
                        else:
                            failure_categories.append("Distance Protection")
                    elif "Underfrequency Load Shedding" in failure:
                        failure_categories.append("Underfrequency Load Shedding")
                    elif failure == "":
                        failure_categories.append("")
                    else:
                        failure_categories.append("Other")
                merged_category = "-".join([i for i in sorted(failure_categories) if i]) # if i removes the empty elements ("")

                mean_proba = sum(branch.proba for branch in branchGroup.branches) / nb_run

                writer.writerow([mean_proba, getElementName(branchGroup.shcLine), branchGroup.end,
                    branchGroup_risk, branchGroup_deviation, branchGroup_cost,
                    branchGroup_cost_deviation, branchGroup_risk/mean_proba/sum(initLoad)*100,
                    branchGroup2_risk, branchGroup2_deviation, branchGroup2_cost,
                    branchGroup2_cost_deviation, branchGroup2_risk/mean_proba/sum(initLoad)*100,
                    branchGroup_risk - branchGroup2_risk, delta_risk_deviation] +
                    delta_risks +
                    [branchGroup_cost - branchGroup2_cost, delta_cost_deviation] +
                    delta_costs +
                    failures + failure_categories + [merged_category])

            # Redo the same with branchGroups2 but only for the branchGroups that have not equivalent in branchGroups(1)
            for branchGroup2 in branchGroups2:
                compareTo = None
                for branchGroup in branchGroups:
                    if set(branchGroup.outOfService) == set(branchGroup2.outOfService) and \
                    branchGroup.shcLine == branchGroup2.shcLine and branchGroup.end == branchGroup2.end:
                        compareTo = branchGroup

                if compareTo == None: # Only for the branchGroups that have not equivalent in branchGroups(1)
                    branchGroup2.branches.sort(key=attrgetter("parameter_set"))

                    branchGroup2_risks = [branch.proba * branch.loadShed for branch in branchGroup2.branches]
                    branchGroup2_risk = sum(branchGroup2_risks)/nb_run
                    branchGroup2_deviation = stdDeviation(branchGroup2_risks)

                    branchGroup2_costs = [branch.proba * loadShedToCost(branch.loadShed) for branch in branchGroup2.branches]
                    branchGroup2_cost = sum(branchGroup2_costs) / nb_run
                    branchGroup2_cost_deviation= stdDeviation(branchGroup2_costs)


                    branchGroup_risk = 0
                    branchGroup_cost = 0
                    branchGroup_deviation = 0
                    branchGroup_cost_deviation = 0


                    delta_risks = [0] * nb_run
                    for branch in branchGroup2.branches:
                        delta_risks[branch.parameter_set] -= branch.proba * branch.loadShed
                    delta_risk_deviation = stdDeviation(delta_risks)


                    delta_costs = [0] * nb_run
                    for branch in branchGroup2.branches:
                        delta_costs[branch.parameter_set] += branch.proba * loadShedToCost(branch.loadShed)
                    delta_cost_deviation = stdDeviation(delta_costs)


                    failures = []
                    for failure in branchGroup2.outOfService:
                        if failure.loc_name[-22:] != "Ph-E Polygonal 3 Delay" and failure.loc_name[-22:] != "Ph-E Polygonal 2 Delay": # Duplicate of Ph-Ph Polygonal * Delay
                            failures.append(getElementName(failure))
                    while len(failures) < 3:
                        failures.append("")

                    failure_categories = []
                    for failure in failures:
                        if "Distance Protection" in failure:
                            if "Delay" in failure:
                                failure_categories.append("Delay")
                            else:
                                failure_categories.append("Distance Protection")
                        elif "Underfrequency Load Shedding" in failure:
                            failure_categories.append("Underfrequency Load Shedding")
                        elif failure == "":
                            failure_categories.append("")
                        else:
                            failure_categories.append("Other")
                    merged_category = "-".join([i for i in sorted(failure_categories) if i]) # if i removes the empty elements ("")

                    mean_proba = sum(branch.proba for branch in branchGroup2.branches)/len(branchGroups[0].branches)

                    writer.writerow([mean_proba, getElementName(branchGroup2.shcLine), branchGroup2.end,
                        branchGroup_risk, branchGroup_deviation, branchGroup_cost,
                        branchGroup_cost_deviation, branchGroup_risk/mean_proba/sum(initLoad)*100,
                        branchGroup2_risk, branchGroup2_deviation, branchGroup2_cost,
                        branchGroup2_cost_deviation, branchGroup2_risk/mean_proba/sum(initLoad)*100,
                        branchGroup_risk - branchGroup2_risk, delta_risk_deviation] +
                        delta_risks +
                        [branchGroup_cost - branchGroup2_cost, delta_cost_deviation] +
                        delta_costs +
                        failures + failure_categories + [merged_category])

    stdError = computeStandardError(current_parameter_set)/computeRisk()*100
    print("\nStandard error on the risk: " + str(stdError))
    filename = results_path + "\\" + dir_name + "\\Std deviation" + ".txt"
    with open(filename, "w") as file:
        file.write(str(stdError))


def importCSV(file):
    """
    To process data in Python instead of Excel
    """
    out = []
    results_path = "C:\\Users\\Mon pc\\Desktop\\"
    filename = results_path + file + ".csv"
    with open(filename, 'r', newline='') as csvfile:
        reader = csv.reader(csvfile, dialect='excel', delimiter=';')
        for row in reader:
            out.append(';'.join(row))
    return out

def runSensitivityStudy(nbRun):
    global start
    start = moduleTime.time()
    global gauss
    gauss = 1
    runMCDET(nbRun)
    RRW = computeImportanceMeasures(1)
    lineImportance = computeLineImportance(1)
    exportToCSV()
    parameterSetsToCSV()

    print("Total execution time : " + str(moduleTime.time()-start))
    filename = results_path + "\\" + dir_name + "\\Computation time" + ".txt"
    with open(filename, "w") as file:
        file.write(str(moduleTime.time()-start))

    return RRW, lineImportance

def runParameterOptimisation(nbRun, recursive=1, Z3failure=1):
    global start
    start = moduleTime.time()
    global gauss
    gauss = 0
    runMCDET(nbRun, recursive, Z3failure)

    parameterSetsToCSV()
    exportToCSV()
    RRW = computeImportanceMeasures(1)
    lineImportance = computeLineImportance(1)

    print("Total execution time : " + str(moduleTime.time()-start))
    filename = results_path + "\\" + dir_name + "\\Computation time" + ".txt"
    with open(filename, "w") as file:
        file.write(str(moduleTime.time()-start))

    for parameter_set in range(len(current_protection_system_branchGroups[0].branches)):
        print(computeRisk(parameter_set))
    return RRW, lineImportance


def compareProtectionSystems(nbRuns = 1, recursive = 1, Z3failure = 1):
    setProtectionSystem(1) # Check that setProtectionSystem(1) is not bugged BEFORE launching the computations
    setBackProtectionSystem() # To be sure it is actually the one used before saveInitialParameters()

    runAll(recursive, Z3failure)
    removeNegligeableGroups()
    saveInitialParameters()

    try:
        setProtectionSystem(1)
        # Rerun the tree with the same parameters and fault positions (cf correlated sampling)
        for line in lines:
            faultLocations = []
            # Copy the fault positions of the branches run with the default protection system
            for branch in root_branches[current_parameter_set]:
                if branch.shcLine == line:
                    if branch.position not in faultLocations:
                        faultLocations.append(branch.position)
            faultLocations.sort()
            buildRoots(line, Z3failure, faultLocations)

        for i in current_protection_system_root_branches[current_parameter_set]:
            i.run(recursive)
        setBackProtectionSystem()

        for i in range(nbRuns-1):  # -1 because one run with default parameter already done
            # Normal run
            randomiseParameters()
            current_protection_system_root_branches.append([])
            print("\nRunning set of branches no: " + str(current_parameter_set))
            runAll(recursive, Z3failure)

            # Run with new protection system
            setProtectionSystem(1)
            current_protection_system_root_branches.append([])
            for line in lines:
                faultLocations = []
                # Copy the fault positions of the branches run with the default protection system
                for branch in root_branches[current_parameter_set]:
                    if branch.shcLine == line:
                        if branch.position not in faultLocations:
                            faultLocations.append(branch.position)
                faultLocations.sort()
                buildRoots(line, Z3failure, faultLocations)

            for i in current_protection_system_root_branches[current_parameter_set]:
                i.run(recursive)
            setBackProtectionSystem()

        setDefaultParameters()
    except Exception as e:
        print(e)
        setBackProtectionSystem()
        setDefaultParameters()
        raise Exception("Simulation stopped safely")
    except:
        print("Unknow exception")
        setBackProtectionSystem()
        setDefaultParameters()
        raise Exception("Simulation stopped safely")

    delta_risks = []
    for i in range(current_parameter_set+1):
        risk1 = 0
        risk2 = 0
        for branchGroup in branchGroups:
            for branch in branchGroup.branches:
                if branch.parameter_set == i:
                    risk1 += branch.proba * branch.loadShed
        for branchGroup in branchGroups2:
            for branch in branchGroup.branches:
                if branch.parameter_set == i:
                    risk2 += branch.proba * branch.loadShed
        delta_risks.append(risk1-risk2)

    print("Delta risk: " + str(sum(delta_risks)/len(delta_risks)))
    if len(delta_risks) > 1:
        print("\nStandard error on the delta risk: " + str(stdDeviation(delta_risks)/sqrt(len(delta_risks)-1)))
    filename = results_path + "\\" + dir_name + "\\Standard error on delta risk" + ".txt"
    with open(filename, "w") as file:
        if len(delta_risks) > 1:
            file.write(str(stdDeviation(delta_risks)/sqrt(len(delta_risks)-1)))
        else:
            file.write(str(0))

    exportToCSV()
    RRW = computeImportanceMeasures(1)
    lineImportance = computeLineImportance(1)
    parameterSetsToCSV()
    filename = results_path + "\\" + dir_name + "\\Computation time" + ".txt"
    with open(filename, "w") as file:
        file.write(str(moduleTime.time()-start))

    return RRW, lineImportance

def setProtectionSystem(state):
    """
    Caution: should not interact with randomiseParameters() => setBackProtectionSystem()
    should always be called before calling randomiseParameters() again
    """
    global current_protection_system_root_branches
    global current_protection_system_branchGroups
    global current_protection_system_skipped_branches
    if state == 0:
        current_protection_system_root_branches = root_branches
        current_protection_system_branchGroups = branchGroups
        current_protection_system_skipped_branches = skipped_branches
    elif state == 1:
        current_protection_system_root_branches = root_branches2
        current_protection_system_branchGroups = branchGroups2
        current_protection_system_skipped_branches = skipped_branches2
    else:
        raise NotImplementedError("Can only compare two protection systems at a time")

    for line in lines:
        for side in range(2):
            cub = line.GetCubicle(side)
            relay = cub.cpRelays
            if state:
                relay.GetContents("Output Logic")[0].aDipset = "30" # PUTT on(high), POTT off


                # cub.GetContents("Line Overcurrent Protection")[0].outserv = 1
                # cub.GetContents("Line Overcurrent Protection")[0].Tpset = 2


                # if side == 0:
                #     cub.GetContents("Interlink")[0].outserv = 0

                # line = app.GetCalcRelevantObjects("Line 23 - 24.ElmLne")[0]
                # line.nbnum = 1

                # X4 = 1 * line.X1
                # R4 = 1 * line.R1 * 2
                # # relay.GetContents("Ph-Ph Polygonal 4")[0].cpXmax = X4 * (1 + random.gauss(0, distance_variance))
                # # relay.GetContents("Ph-E Polygonal 4")[0].cpXmax = X4 * (1 + random.gauss(0, distance_variance))
                # # relay.GetContents("Ph-Ph Polygonal 4")[0].cpRmax = R4 * (1 + random.gauss(0, distance_variance))
                # # relay.GetContents("Ph-E Polygonal 4")[0].cpRmax = R4 * (1 + random.gauss(0, distance_variance))
                # # relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].Tdelay = 0.04 + random.gauss(0, delay_flat_variance)
                # # relay.GetContents("Ph-E Polygonal 4 Delay")[0].Tdelay = 0.04 + random.gauss(0, delay_flat_variance)
                # relay.GetContents("Ph-Ph Polygonal 4")[0].cpXmax = X4*0.99
                # relay.GetContents("Ph-E Polygonal 4")[0].cpXmax = X4*0.99
                # relay.GetContents("Ph-Ph Polygonal 4")[0].cpRmax = R4*0.99
                # relay.GetContents("Ph-E Polygonal 4")[0].cpRmax = R4*0.99
                # relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].Tdelay = 0.04
                # relay.GetContents("Ph-E Polygonal 4 Delay")[0].Tdelay = 0.04

                # relay.GetContents("Ph-Ph Polygonal 4")[0].outserv = 0
                # relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].outserv = 0
                # relay.GetContents("Ph-E Polygonal 4")[0].outserv = 0
                # relay.GetContents("Ph-E Polygonal 4 Delay")[0].outserv = 0
            else:
                relay.GetContents("Output Logic")[0].aDipset = "00" # PUTT off, POTT off


                # cub.GetContents("Line Overcurrent Protection")[0].outserv = 0
                # cub.GetContents("Line Overcurrent Protection")[0].Tpset = 2


                # if side == 0:
                #     cub.GetContents("Interlink")[0].outserv = 1

                # line = app.GetCalcRelevantObjects("Line 23 - 24.ElmLne")[0]
                # line.nbnum = 2 # Double the capacity of the line

                # relay.GetContents("Ph-Ph Polygonal 4")[0].outserv = 1
                # relay.GetContents("Ph-Ph Polygonal 4 Delay")[0].outserv = 1
                # relay.GetContents("Ph-E Polygonal 4")[0].outserv = 1
                # relay.GetContents("Ph-E Polygonal 4 Delay")[0].outserv = 1


def setBackProtectionSystem():
    setProtectionSystem(0)

if IEEE9:
    # runAll(0,0)
    # exportToCSV()
    compareProtectionSystems(3,1,1)
    # runMCDET(10)
    # (RRW, lineImportance) = runParameterOptimisation(5)
    # (RRW, lineImportance) = runSensitivityStudy(10)
else:
    runMCDET(50, 0, 0)
    parameterSetsToCSV()
    # runAll(1,1)
    exportToCSV()
    RRW = computeImportanceMeasures(1)
    lineImportance = computeLineImportance(1)
    endImportance = computeEndImportance(1)
    filename = results_path + "\\" + dir_name + "\\Computation time" + ".txt"
    with open(filename, "w") as file:
        file.write(str(moduleTime.time()-start))

    # (RRW, lineImportance) = runParameterOptimisation(5)
    # (RRW, lineImportance) = runSensitivityStudy(5)
    # compareProtectionSystems(1,1,1)

    # !!!!!!!!!! Before launch, check gauss, randomisePosition, frequency_cutoff and runAll(1,1) or runAll(0,0)
