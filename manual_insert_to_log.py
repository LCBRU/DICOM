

import numpy as np

to_log = [',,1,2021-09-14 09:14:00.000000,2021-09-14 09:37:00.000000,']


with open("C:\\briccs_ct\\results.csv", "ab") as f:
    np.savetxt(f, (to_log), fmt='%s', delimiter=' ')