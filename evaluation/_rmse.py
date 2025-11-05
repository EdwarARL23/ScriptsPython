"""
Function that implements the root mean square error (RMSE) metric.
"""


import numpy as np
from evaluation._ancillary_functions import _process_inputs_for_metrics


def RMSE(p_real, p_pred):

    
    # Checking if inputs are compatible
    p_real, p_pred = _process_inputs_for_metrics(p_real, p_pred)

    return np.sqrt(np.mean((p_real - p_pred)**2))
