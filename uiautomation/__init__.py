from __future__ import absolute_import

# com init with multithread mode to enable callback function
import sys
sys.coinit_flags = 0x0

from .version import VERSION
from .uiautomation import *
