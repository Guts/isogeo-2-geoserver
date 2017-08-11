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
import logging

# Python 3 backported
from collections import OrderedDict

# 3rd party libraries
from geoserver.catalog import Catalog
from geoserver.resource import Coverage, FeatureType

# ############################################################################
# ######### Classes #############
# ###############################


class ReadGeoServer():
    def __init__(self, gs_axx, dico_gs, tipo, txt=''):
        """Use OGR functions to extract basic informations about geoserver.

        gs_axx = tuple like {url of a geoserver, user, password)
        dico_gs = dictionary to store
        tipo = format
        text = dictionary of text in the selected language
        """
        # connection
        cat = Catalog(gs_axx[0], gs_axx[1], gs_axx[2],
                      disable_ssl_certificate_validation=gs_axx[3])
        # print(dir(cat))



        # -- WORKSPACES -------------------------------------------------------
        workspaces = cat.get_workspaces()
        for wk in workspaces:
            # print(wk.name, wk.enabled, wk.resource_type, wk.wmsstore_url)
            dico_gs[wk.name] = wk.href, {}
        # print(dir(wk))

        # -- STORES -----------------------------------------------------------
        # stores = cat.get_stores()
        # for st in stores:
        #     # print(st.name, st.enabled, st.href, st.resource_type)
        #     if hasattr(st, 'url'):
        #         url = st.url
        #     elif hasattr(st, 'resource_url'):
        #         url = st.resource_url
        #     else:
        #         print(dir(st))

        #     dico_gs.get(st.workspace.name)[1][st.name] = {"ds_type": st.type,
        #                                                   "ds_url": url}

        # print(dir(st))
        # print(st.url)
        # print(st.workspace)
        # print(dir(st.workspace))
        # print(st.workspace.name)

        # -- LAYERS -----------------------------------------------------------
        # resources_target = cat.get_resources(workspace='ayants-droits')
        layers = cat.get_layers()
        logging.info("{} layers found".format(len(layers)))
        dico_layers = OrderedDict()
        for layer in layers:
            # print(layer.resource_type)
            lyr_title = layer.resource.title
            lyr_name = layer.name
            lyr_wkspace = layer.resource._workspace.name
            if type(layer.resource) is Coverage:
                lyr_type = "coverage"
            elif type(layer.resource) is FeatureType:
                lyr_type = "vector"
            else:
                lyr_type = type(layer.resource)

            # a log handshake
            logging.info("{} | {} | {} | {}".format(layers.index(layer),
                                                    lyr_type,
                                                    lyr_name,
                                                    lyr_title))

            # # METADATA LINKS #
            # download link
            md_link_dl = "{0}/geoserver/{1}/ows?request=GetFeature"\
                         "&service=WFS&typeName={1}%3A{2}&version=2.0.0"\
                         "&outputFormat=SHAPE-ZIP"\
                         .format(url_base,
                                 lyr_wkspace,
                                 lyr_name)

            # mapfish links
            md_link_mapfish_wms = "{0}/mapfishapp/?layername={1}"\
                                  "&owstype=WMSLayer&owsurl={0}/"\
                                  "geoserver/{2}/ows"\
                                  .format(url_base,
                                          lyr_name,
                                          lyr_wkspace)

            md_link_mapfish_wfs = "{0}/mapfishapp/?layername={1}"\
                                  "&owstype=WFSLayer&owsurl={0}/"\
                                  "geoserver/{2}/ows"\
                                  .format(url_base,
                                          lyr_name,
                                          lyr_wkspace)

            md_link_mapfish_wcs = "{0}/mapfishapp/?cache=PreferNetwork"\
                                  "&crs=EPSG:2154&format=GeoTIFF"\
                                  "&identifier={1}:{2}"\
                                  "&url={0}/geoserver/ows?"\
                                  .format(url_base,
                                          lyr_wkspace,
                                          lyr_name)

            # OC links
            md_link_oc_wms = "{0}/geoserver/{1}/wms?layers={1}:{2}"\
                             .format(url_base,
                                     lyr_wkspace,
                                     lyr_name)
            md_link_oc_wfs = "{0}/geoserver/{1}/ows?typeName={1}:{2}"\
                             .format(url_base,
                                     lyr_wkspace,
                                     lyr_name)

            md_link_oc_wcs = "{0}/geoserver/{1}/ows?typeName={1}:{2}"\
                             .format(url_base,
                                     lyr_wkspace,
                                     lyr_name)

            # CSW Querier links
            md_link_csw_wms = "{0}/geoserver/ows?service=wms&version=1.3.0"\
                              "&request=GetCapabilities".format(url_base)

            md_link_csw_wfs = "{0}/geoserver/ows?service=wfs&version=2.0.0"\
                              "&request=GetCapabilities".format(url_base)

            # # # SERVICE LINKS #
            # GeoServer Edit links
            gs_link_edit = "{}/geoserver/web/?wicket:bookmarkablePage="\
                           ":org.geoserver.web.data.resource."\
                           "ResourceConfigurationPage"\
                           "&name={}"\
                           "&wsName={}".format(url_base,
                                               lyr_wkspace,
                                               lyr_name)

            # Metadata links (service => metadata)
            if is_uuid(dict_match_gs_md.get(lyr_name)):
                # HTML metadata
                md_uuid_pure = dict_match_gs_md.get(lyr_name)
                srv_link_html = "{}/portail/geocatalogue?uuid={}"\
                                .format(url_base, md_uuid_pure)

                # XML metadata
                md_uuid_formatted = "{}-{}-{}-{}-{}".format(md_uuid_pure[:8],
                                                            md_uuid_pure[8:12],
                                                            md_uuid_pure[12:16],
                                                            md_uuid_pure[16:20],
                                                            md_uuid_pure[20:])
                srv_link_xml = "http://services.api.isogeo.com/ows/s/"\
                               "{1}/{2}?"\
                               "service=CSW&version=2.0.2&request=GetRecordById"\
                               "&id=urn:isogeo:metadata:uuid:{0}&"\
                               "elementsetname=full&outputSchema="\
                               "http://www.isotc211.org/2005/gmd"\
                               .format(md_uuid_formatted,
                                       csw_share_id,
                                       csw_share_token)
                # add to GeoServer layer
                rzourc = cat.get_resource(lyr_name,
                                          store=layer.resource._store.name)
                rzourc.metadata_links = [('text/html', 'ISO19115:2003', srv_link_html),
                                         ('text/xml', 'ISO19115:2003', srv_link_xml),
                                         ('text/html', 'TC211', srv_link_html),
                                         ('text/xml', 'TC211', srv_link_xml)]
                # rzourc.metadata_links.append(('text/html', 'other', 'hohoho'))
                cat.save(rzourc)

            else:
                logging.info("Service without metadata: {} ({})".format(lyr_name,
                                                                        dict_match_gs_md.get(lyr_name)))
                pass

            dico_layers[layer.name] = {"title": lyr_title,
                                       "workspace": lyr_wkspace,
                                       "store_name": layer.resource._store.name,
                                       "store_type": layer.resource._store.type,
                                       "lyr_type": lyr_type,
                                       "md_link_dl": md_link_dl,
                                       "md_link_mapfish": md_link_mapfish_wms,
                                       "md_link_mapfish_wms": md_link_mapfish_wms,
                                       "md_link_mapfish_wfs": md_link_mapfish_wfs,
                                       "md_link_mapfish_wcs": md_link_mapfish_wcs,
                                       "md_link_oc_wms": md_link_oc_wms,
                                       "md_link_oc_wfs": md_link_oc_wfs,
                                       "md_link_oc_wcs": md_link_oc_wcs,
                                       "md_link_csw_wms": md_link_csw_wms,
                                       "md_link_csw_wfs": md_link_csw_wfs,
                                       "gs_link_edit": gs_link_edit,
                                       "srv_link_html": srv_link_html,
                                       "srv_link_xml": srv_link_xml,
                                       "md_id_matching": md_uuid_pure
                                       }

            # mem clean up
            # del dico_layer

        # print(dico_gs.get(layer.resource._workspace.name)[1][layer.resource._store.name])
        # print(dir(layer.resource))
        # print(dir(layer.resource.writers))
        # print(dir(layer.resource.writers.get("metadataLinks").func_dict))
        # print(layer.resource.metadata)
        # print(layer.resource.metadata_links)

        dico_gs["layers"] = dico_layers