---
layout: page
permalink: publications/powertech2023
---

## Abstract

Cascading outages in power systems are complex phenomena that are difficult to model and not yet fully understood. When attempting to simulate cascading outages, close attention should thus be paid to modelling uncertainties. This is especially true for fast cascading outages (i.e. cascading outages that are driven by electromechanical transients) during which many protection systems can operate in close succession. Indeed, small variations in the timing of a protection operation can vastly impact the cascade propagation and its consequences. A classical approach to account for modelling uncertainties is to perform Monte Carlo (MC) simulations which is computationally challenging. To alleviate this challenge, we propose an indicator based on sequences of tripping events to predict which contingencies are more sensitive to uncertainties and which are less. MC simulations can be avoided for the latter, saving computation time. The performance of our indicator is demonstrated using a modified version of the IEEE 39-bus test system. We show that analysing the sequences of tripping events leads to better performing indicators than looking only at the consequences (e.g. load shedding) of cascades. We also show how the proposed indicator can be used to speed up a simplified probabilistic security assessment.

[Download full text here](https://difusion.ulb.ac.be/vufind/Record/ULB-DIPOT:oai:dipot.ulb.ac.be:2013/368959/Holdings)

## Citation

{% raw %}
```
@INPROCEEDINGS{ISGT2023_Protections,
    author={Sabot, Frédéric and Labeau, Pierre-Etienne and Henneaux, Pierre},
    booktitle={2023 IEEE PES Innovative Smart Grid Technologies Europe (ISGT EUROPE)},
    title={Handling protection-related uncertainties in simulations of fast cascading outages},
    year={2023},
    month = 10,
    doi={10.1109/ISGTEUROPE56780.2023.10408629}
}
```
{% endraw %}

## Supplementary data

**Note:** The links below provide with the data that were used to write the above publication. For more up-to-date versions, please refer to the master branch of the linked repositories.

[Modified IEEE39](https://github.com/FredericSabot/dynawo/tree/15_PowerTech2023/examples/DynaSwing/IEEE39/IEEE39_Fault). Please refer to the paper for more details about the modifications

[Scripts](https://github.com/FredericSabot/dynawo-algorithms/tree/4_PowerTech2023/Scripts/SecurityAssessment). The MC simulations were performed with the batch_N-2_cluster.sh script. Computation of the indicators was performed with the batch_N-2.sh and batch_N-2_timeline.sh scripts. For more details, please refer to the documentation of the .py files or contact me.

