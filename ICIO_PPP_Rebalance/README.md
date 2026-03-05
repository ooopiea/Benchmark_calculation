# ICIO-PPP-GRAS
This project is based on ICIO published by OECD. We use PPP by WB to modifed the ICIO, and apply GRAS to rebalance the table.

## data files:

### Original data:
(1) **2022_SML.xlsx**: from OECD ICIO
(2) **WB-ER.xlsx**: from WB
(3) **WB-PPP.xlsx**: from WB
(4) **PLI.xlsx**: preconditioned from **WB-ER.xlsx** and **WB-PPP.xlsx**

### Processed data:
(1) **modified_ICIO.xlsx**: generated from icio_PLI_transform.py
    - This table is unbalanced ICIO modified by PLI.
(2) **balanced_ICIO.xlsx**: generated from gras_icio_rebalance.py
    - This table is balanced ICIO (USD PPP).
(3) **rebalanced_diff.xlsx**: generated from rebalance_verification.py
    - This table demonstrates difference between **modified_ICIO** and **balanced_ICIO.xlsx** from ratio and absolute difference.
(4) **rebalanced_diff_large_diff.xlsx**: generated from rebalance_verification.py
    - This table selects data points with large difference in a long table.
    