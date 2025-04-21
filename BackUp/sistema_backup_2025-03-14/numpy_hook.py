import os
import sys
import numpy as np

# Force numpy to load from the bundled package
numpy_path = os.path.dirname(np.__file__)
if hasattr(sys, '_MEIPASS'):
    os.environ['NUMPY_LIBS'] = os.path.join(sys._MEIPASS, 'numpy.libs')
    if numpy_path.startswith(sys._MEIPASS):
        sys.path.insert(0, os.path.join(sys._MEIPASS, 'numpy'))
