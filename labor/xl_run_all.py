"""
<copyright>
Copyright (c) 2025, Janusch Rentenatus. This program and the accompanying materials are made available under the
terms of the Apache License v2.0 which accompanies this distribution, and is available at
https://github.com/Rentenatus/py_yahtzee?tab=Apache-2.0-1-ov-file#readme
</copyright>
"""

from labor.xl_step01_var import Step01
from labor.xl_step02_sign import Step02
from labor.xl_step03_code import Step03

if __name__ == "__main__":
    step = Step01()
    step.run()
    step = Step02()
    step.run()
    step = Step03()
    step.run()