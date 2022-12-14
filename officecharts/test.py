import pandas as pd
import numpy as np
import officecharts.container

data = pd.DataFrame({"value": np.cumprod(1 + np.append(0, np.random.normal(0.005, 0.008, 11))) - 1},
                    index=pd.date_range('2022-01-31', '2022-12-31', freq='M'))

officecharts.container.clipboard_container(data)
