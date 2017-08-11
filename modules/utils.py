# -*- coding: UTF-8 -*-
#!/usr/bin/env python
from __future__ import (absolute_import, print_function, unicode_literals)
# ----------------------------------------------------------------------------
# Name:         Geodata Explorer
# Purpose:      Explore directory structure and list files and folders
#               with geospatial data
#
# Author:       Julien Moura (@geojulien)
#
# Python:       2.7.x
#
# Licence:      GPL 3
# ----------------------------------------------------------------------------

# ############################################################################
# ######## Libraries #############
# ################################

# Standard library
from uuid import UUID

# ############################################################################
# ######### Classes #############
# ###############################


class Utils(object):
    def __init__(self):
        u"""DicoGIS specific utilities"""
        super(Utils, self).__init__()

    def tunning_worksheets(li_worksheets):
        """CLEAN UP & TUNNING worksheets list."""
        for sheet in li_worksheets:
            # Freezing panes
            c_freezed = sheet['B2']
            sheet.freeze_panes = c_freezed

            # Print properties
            sheet.print_options.horizontalCentered = True
            sheet.print_options.verticalCentered = True
            sheet.page_setup.fitToWidth = 1
            sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE

            # Others properties
            wsprops = sheet.sheet_properties
            wsprops.filterMode = True

            # enable filters
            sheet.auto_filter.ref = str("A1:{}{}")\
                                    .format(get_column_letter(sheet.max_column),
                                                sheet.max_row)
        pass

    def is_uuid(uuid_string, version=4):
        """Si uuid_string est un code hex valide mais pas un uuid valid,
        UUID() va quand même le convertir en uuid valide. Pour se prémunir
        de ce problème, on check la version original (sans les tirets) avec
            # le code hex généré qui doivent être les mêmes.
        """
        try:
            uid = UUID(str(uuid_string), version=version)
            return uid.hex == str(uuid_string).replace('-', '')
        except ValueError as e:
            logger.error("uuid ValueError. {} ({})  -- {}".format(type(uuid_string),
                                                                  uuid_string,
                                                                  e))
            return False
        except TypeError:
            logger.error("uuid must be a string. Not: {} ({})".format(type(uuid_string),
                                                                      uuid_string))
            return False
