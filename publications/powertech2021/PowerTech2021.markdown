---
layout: page
title: MCDET as a Tool for Probabilistic Dynamic Security Assessment of Transmission Systems
permalink: publications/powertech2021
---

## Abstract

In the field of power systems, a fast cascade is defined as the part of a cascading outage that is driven by electromechanical transients. During a fast cascade, the time interval between two protection system operations can be as small as a few milliseconds. The order of events and the global evolution of the cascade are thus very sensitive to measurement errors and epistemic uncertainties. The stochastic failures of protection devices are another source of uncertainty. Therefore, a dynamic probabilistic security assessment is necessary to fully understand and better mitigate fast cascading outages. Dynamic security assessment methodologies have already been applied to the study of the fast cascade, but they require the use of a custom-built simulator and their results are difficult to interpret and to exploit. In this work, we propose to apply the powertech2021 (Monte Carlo Dynamic Event Tree) methodology to the probabilistic dynamic security assessment of transmission systems. The methodology is implemented using the DIgSILENT PowerFactory simulator and tested on the IEEE 39-bus system. Importance measures are used to identify assets whose maintenance or replacement should be prioritised, and sensitivities are used to identify inadequate or sensitive protection settings.

[Download full text here](https://difusion.ulb.ac.be/vufind/Record/ULB-DIPOT:oai:dipot.ulb.ac.be:2013/331914/Holdings)

## Citation

```bibtex
@INPROCEEDINGS{MCDETasTool,
    author={Sabot, Frédéric and Henneaux, Pierre and Labeau, Pierre-Etienne},
    booktitle={2021 IEEE Madrid PowerTech},
    title={MCDET as a Tool for Probabilistic Dynamic Security Assessment of Transmission Systems},
    year={2021},
    month = 6,
    doi={10.1109/PowerTech46648.2021.9494758}
}
```

## Supplementary data

[IEEE 39 test system in PowerFactory 2019 SP4 format](https://fredericsabot.github.io//publications/powertech2021/39 Bus New England System.pfd)

[Python script to add protections and/or change their settings](https://fredericsabot.github.io//publications/powertech2021/EditProtections.py)

[Python script to generate dynamic event trees, solve them using PowerFactory and output results in CSV files](https://fredericsabot.github.io//publications/powertech2021/main.py)

