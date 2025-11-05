"""
Function that implements the symmetric mean absolute percentage error (sMAPE) metric.
"""

import numpy as np
from evaluation._ancillary_functions import _process_inputs_for_metrics


def sMAPE(p_real, p_pred):


    # Checking if inputs are compatible
    p_real, p_pred = _process_inputs_for_metrics(p_real, p_pred)

    return np.mean(np.abs(p_real - p_pred) / ((np.abs(p_real) + np.abs(p_pred)) / 2))
