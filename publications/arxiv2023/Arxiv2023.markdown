---
layout: page
permalink: publications/arxiv2023
---

## Abstract

With the transition towards a smart grid, Information and Communications Technology (ICT) infrastructures play a growing role in the operation of transmission systems. Cyber-physical systems are usually studied using co-simulation. The latter allows to leverage existing simulators and models for both power systems and communication networks. A major drawback of co-simulation is however the computation time. Indeed, simulators have to be frequently paused in order to stay synchronised and to exchange information. This is especially true for large systems for which lots of interactions between the two domains occur. We thus propose a self-consistent simulation approach as an alternative to co-simulation. We compare the two approaches on the IEEE 39-bus test system equipped with an all-PMU state estimator. We show that our approach can reach the same accuracy as co-simulation, while using drastically less computer resources.

[Download full text here](https://arxiv.org/abs/2311.14467)

## Citation

```bibtex
@misc{sabot2023cosim,
      title={Towards an Efficient Simulation Approach for Transmission Systems with ICT Infrastructures},
      author={Frédéric Sabot and Pierre-Etienne Labeau and Jean-Michel Dricot and Pierre Henneaux},
      year={2023},
      eprint={2311.14467},
      archivePrefix={arXiv},
      primaryClass={eess.SY}
}
```

## Supplementary data

**Note:** The links below provide with the data that were used to write the above publication. If you need a more up-to-date version of Dynawo, you should rebase FredericSabot/dynawo/9_Cosim onto the master branch of dynawo/dynawo.

If you need a more up-to-date version of Helics or ns-3, I will assume you know what you are doing. Also, it can be noted that there is currently a co-simulation platform linking Dynawo, ns-3 and omnet++ currently in development in the framework of the CYPRESS project (https://cypress-project.be/) that should be made available on [https://github.com/Adrien-Leerschool/CYPRESS_Co-Simulation_Platform](https://github.com/Adrien-Leerschool/CYPRESS_Co-Simulation_Platform)

## Introduction

I interfaced Dynawo as a so-called “ValueFederate”, i.e. a federate that publishes and reads some values. This is as opposed to a MessageFederate that communicates through... messages. This indeed makes more sense for a physical simulator. On ns-3 side, I thus added a PMU class that reads those values at a given frequency and sends them as messages in the communication network (messages are sent to PDCs, then a SPDC (control centre) as in the GECO test case). (This means that only ns-3 knows the sampling frequency of the PMUs which makes more sense as it is more of a “cyber parameter”.)

### Time management

There is two main ways to manage time in a co-simulation: time-stepped and event based (see previous presentations). We use the time-stepped approach for the following reasons:

* Dynawo cannot say in advance when will be its next step/event. This is because the next event is determined by zero-crossings/time step variations which are known only after solving the DAEs.
* As an event-based simulator, ns-3 naturally knows its next time step. However, it does not know its next “interesting” (from the co-simulation point of view). To give an example, if Dynawo sends a message to ns-3, message that goes through a few routers, then is sent back to Dynawo, the next “interesting” event, is when the message goes back to Dynawo. However, in the event queue of ns-3, initially the only event is “Dynawo message sent to first router”. Then, when ns-3 process this event, the “message sent from first to second router” event is created. Ns-3 thus have to go through all these events (stopping the co-simulation each time) before arriving at the “interesting” event. This would be very inefficient especially for large systems.

So, in the example, Dynawo stops every 1ms to read and write values. Ns-3 is event based (but Helics allows to continue without stopping for synchronising except every 1ms when Dynawo asks). The PMUs in ns-3 only read the values published by Dynawo every 20ms, but Dynawo still reads/write every 1ms to be able to react quickly when ns-3 sends it a message (e.g. control action decided from the PMU measurements). (It is a bit inefficient for Dynawo to publish every 1ms when ns-3 only reads every 20ms, but not that much. It could be possible to only sends those measurements every 20ms (while still stopping Dynawo every 1ms) but I don’t expect a significant performance gain. Also, there is a functionality in Helics to only publish data when it significantly changes.)

## Installation

First, install Helics v3.x following the instructions on https://docs.helics.org/en/latest/user-guide/installation/index.html. If building from source, make sure to turn on the CXX interface.

Ns-3

Install ns-3 and the helics add-on following the instructions on https://github.com/FredericSabot/helics-ns3, but replace
```
git clone https://github.com/GMLC-TDC/helics-ns3 contrib/helics
```
by
```
git clone https://github.com/FredericSabot/helics-ns3.git contrib/helics
```

It is not mandatory, but I personally use the following ./waf configs:

```
./waf configure --with-helics=/usr/local/include/helics --enable-examples --enable-tests --out=build/debug --build-profile=debug
```
or
```
./waf configure --with-helics=/usr/local/include/helics --enable-examples --enable-tests --out=build/opti --build-profile=optimized
```

Dynawo

Build from source following the instructions on https://dynawo.github.io/install/, but replace
```
git clone https://github.com/dynawo/dynawo.git dynawo
```
by
```
git clone https://github.com/FredericSabot/dynawo.git dynawo
git checkout 9_Cosim
```

Normally, Cmake should be able to find Helics. If not, you can tell the install path of helics in dynawo/sources/Modeler/ModelManager/CMakeLists.txt line 31, and dynawo/sources/ModelicaCompiler/compileCppModelicaModelInDynamicLib.cmake.in line 87 (replace /usr/local/lib with the correct path).


## Usage

Launching simulations

The easiest way to launch simulations is via the [helics-cli](https://github.com/GMLC-TDC/helics-cli), simply execute
```
helics run --path ns3-dynawo-full.json
```
with the content of ns3-dynawo-full.json being

```
{
  "federates": [
    {
      "directory": ".",
      "exec": "helics_broker -f 2 --loglevel=Debug",
      "host": "localhost",
      "name": "broker"
    },
    {
      "directory": "/home/fsabot/Desktop/dynawo_new",
      "exec": "./myEnvDynawo.sh jobs examples/DynaSwing/IEEE39/IEEE39_Cosim/IEEE39.jobs",
      "host": "localhost",
      "name": "Dynawo"
    },
    {
      "directory": "/home/fsabot/Desktop/Cosim/ns-3-dev",
      "exec": "./waf --run ieee39",
      "host": "localhost",
      "name": "ns-3"
    }
  ],
  "name": "ns3-Dynawo"
}
```

with path updated based on where you installed the dynawo and ns-3.

## Writing simulations

In Dynawo, the co-simulation interface is done through a model in the CosimulationAutomaton library (the id of the model must be CosimInterface). This model has a vector of inputs and a vector of ouputs that are read/written at a given period.

In order to ease the use of Dynawo for co-simulations, I wrote a python script

(DynawoHelicsJsonParser.py available in the root directory of Dynawo)

to translate the Helics configuration file (.json) of the co-simulation (see e.g. examples/DynaSwing/IEEE39/IEEE39_Cosim/Dynawo.json) into the files needed by Dynawo. More precisely, it creates a dyd.out file to be added inside the .dyd file of the simulation, and a par.out to be added inside the .par file of the simulation. Note that in addition to the usual information that should be contained in the Helics config, you should also add an “info” tag for each publication/subscription that tell where to find the value in Dynawo. E.g.
  "publications": [
    {
      "key": "U_1",
      "info": "_BUS____1_TN_Upu_value@NETWORK"
    }
]

Means that the publication U_1 is the variable _BUS____1_TN_Upu_value of the model NETWORK.

In ns-3, it is less automated, but basically, the applications that communicate with the co-simulation (directly or indirectly) should inherit the class helics-application.

