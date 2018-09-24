"""High performance Excel IO tools for pandas DataFrame

See Also
--------
xp.read
    read DataFrame from Excel file
xp.save
    save DataFrame to Excel file
xp.view
    view DataFrame in Excel app
"""

from . import core
from . import utils
from .core import read, save, view, xw

__version__ = '0.1.0'
