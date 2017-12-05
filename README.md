Isogeo to geoserver
===================

## Procédure

1. Télécharger le fichier de correspondance au format .xlsx
2. Le renommer pour enlever tous les caractères spéciaux
3. Le nettoyer au besoin pour que tout soit dans le premier onglet du tableur
4. Remplir le fichier settings.ini

### Structure fichier xlsx

| GS_WORKSPACE | GS_DATASTORE_NAME | GS_DATASTORE_TYPE | GS_SOURCE_TYPE | GS_NOM | GS_TITRE	MD | ISOGEO_UUID |
| :--: | :-- | :--: | :--: | :--: | :--: | :--: |
| GeoServer Workspace name | GeoServer datastore name | GeoServer datastore type | GeoServer source type | GeoServer layer name | GeoServer layer title | Isogeo resource ID |

