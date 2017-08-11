# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ------------------------------------------------------------------------------
# Name:         Isogeo to GeoServer 2.8+
# Purpose:      Synchronize Isogeo and GeoServer
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
# Created:      14/08/2016
# Updated:      28/08/2017
# ------------------------------------------------------------------------------

# ##############################################################################
# ########## Libraries #############
# ##################################

# Standard
import logging
from logging.handlers import RotatingFileHandler

# Python 3 backported


# 3rd party


# ##############################################################################
# ############ Globals ############
# #################################

# LOG
logger = logging.getLogger("isogeo2geoserver")
logging.captureWarnings(True)
logger.setLevel(logging.INFO)  # all errors will be get
log_form = logging.Formatter("%(asctime)s || %(levelname)s || %(module)s || %(message)s")
logfile = RotatingFileHandler("LOG_i2gs.log", "a", 5000000, 1)
logfile.setLevel(logging.INFO)
logfile.setFormatter(log_form)
logger.addHandler(logfile)

# ##############################################################################
# ########## Classes ###############
# ##################################


class IsogeoToGeoServer(object):
    """IsogeoToDocx class."""

    def __init__(self, lang="FR"):
        """Instanciating class."""
        super(IsogeoToGeoServer, self).__init__()

        # ------------ VARIABLES ---------------------



    def function(self):
        # method ending
        return

    def function(self):
        # method ending
        return

    def function(self):
        # method ending
        return

    def function(self):
        # method ending
        return

    def function(self):
        # method ending
        return

    def function(self):
        # method ending
        return

# ############################################################################
# ##### Stand alone program ########
# ##################################

if __name__ == '__main__':
    u"""Standalone execution for tests."""
    # ------------ Specific imports ---------------------
    i2gs = IsogeoToGeoServer()
    print(dir(i2gs))
