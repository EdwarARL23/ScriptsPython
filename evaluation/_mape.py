"""
Function that implements the mean absolute percentage error (MAPE) metric.
"""

import numpy as np
from evaluation._ancillary_functions import _process_inputs_for_metrics


def MAPE(p_real, p_pred, noNaN=False):

    # Checking if inputs are compatible
    p_real, p_pred = _process_inputs_for_metrics(p_real, p_pred)

    # Computing MAPE at every time point
    mape = np.abs(p_real - p_pred) / np.abs(p_real)

    # Eliminating NaN values if requested and averaging
    if noNaN:
        mape = np.mean(mape[np.isfinite(mape)])
    else:
        mape = np.mean(mape)

    return mape
