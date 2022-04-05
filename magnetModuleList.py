#!/usr/bin/env python
""" Script reads the magnet module assembly information and records it
in a dictionary.   The dictionary is then jsonified and put into a file """

from CdbApiFactory import CdbApiFactory
#import click
import json
#from rich import print

CDBItemID = {}
CDBItemID['DLMA'] = 110353
CDBItemID['DLMB'] = 110354
CDBItemID['FODO'] = 110371
CDBItemID['QMQA'] = 110369
CDBItemID['QMQB'] = 110370   

MagnetOrder = {}
MagnetOrder['DLMA'] = ["Q1","FC1","Q2","M1","Q3","S1","Q4","S2","Q5","FC2","S3"]
MagnetOrder['DLMB'] = ["S3","FC2","Q5","S2","Q4","S1","Q3","M1","Q2","FC1","Q1"]
MagnetOrder['FODO'] = ["A:M3","A:Q8","A:M4","B:Q8","B:M3"]
MagnetOrder['QMQA'] = ["Q6","M2","Q7"]
MagnetOrder['QMQB'] = ["Q7","M2","Q6"]


MagnetPrefix = {}
MagnetPrefix['DLMA'] = "A:"
MagnetPrefix['DLMB'] = "B:"
MagnetPrefix['FODO'] = ""
MagnetPrefix['QMQA'] = "A:"
MagnetPrefix['QMQB'] = "B:"

 
# @click.command()
def get_modules():
    url_prefix = "https://cdb.aps.anl.gov/cdb/views/item/view?id="
    apiFactory =  CdbApiFactory("https://cdb.aps.anl.gov/cdb")
    itemApi = apiFactory.getItemApi()
    magnet_module_assignments = {}
    for magnet_module in CDBItemID.keys():
        for inv_item in itemApi.get_items_derived_from_item_by_item_id(CDBItemID[magnet_module]):
            item_hierarchyOBJ = itemApi.get_item_hierarchy_by_id(inv_item.id)
            module_assembly_assignments = {}
            for assembly_item in item_hierarchyOBJ.child_items:
                for mag_index in range(len(MagnetOrder[magnet_module])):
                    element_name = MagnetPrefix[magnet_module] + MagnetOrder[magnet_module][mag_index]
                    if element_name == assembly_item.derived_element_name and assembly_item.item != None:
                        module_data = {}
                        module_data['order'] = mag_index
                        module_data['label'] = MagnetOrder[magnet_module][mag_index]
                        module_data['name'] = assembly_item.derived_item.name
                        module_data['url'] = url_prefix + str(assembly_item.item.id)
                        module_data['serial'] = assembly_item.item.name
                        module_assembly_assignments[assembly_item.derived_element_name] = module_data
            magnet_module_assignments[inv_item.name] = module_assembly_assignments
#    print(json.dumps(magnet_module_assignments,indent=3))
    return(magnet_module_assignments)

MAGNETMODULES = get_modules()

def read_data_test(module):
    return(MAGNETMODULES[module])


if __name__ == "__main__":
    print(json.dumps(MAGNETMODULES,indent=3))    

          
