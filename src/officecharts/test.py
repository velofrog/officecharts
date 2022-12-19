import pandas as pd
import numpy as np
# from .container import create_chart

data = pd.DataFrame({"series 1": np.cumprod(1 + np.append(0, np.random.normal(0.005, 0.008, 11))) - 1,
                     "series 2": np.cumprod(1 + np.append(0, np.random.normal(0.005, 0.008, 11))) - 1,
                     "series 3": np.cumprod(1 + np.append(0, np.random.normal(0.005, 0.008, 11))) - 1},
                    index=pd.date_range('2022-01-31', '2022-12-31', freq='M'))


# import matplotlib.colors
# import matplotlib.pyplot as plt
# import numpy as np
#
# matplotlib.colors.to_hex(matplotlib.colormaps.get_cmap("tab10")(0))
# x = np.linspace(0, 10, 100)
# for i in range(0, 10):
#     plt.plot(x, x * i / 10.0 + 0, color=matplotlib.colormaps.get_cmap("tab10")(i))
