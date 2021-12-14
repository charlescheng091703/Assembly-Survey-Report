#!/usr/bin/env python
""" Box Dead Drop:
Module sets up a BoxDeadDrop class which is used to send a 
serialized data structurein a file to box.com and then retrievs that data 
through the Box Api """



from boxsdk import Client, JWTAuth
import click
import json

DEADDROP_ID = 890210016453

class BoxDeadDrop():
    """ Creates a Dead Drop Class for Box.  Class manages a particular 
        Box File ID as a JSON container to use as a conduit for maintaining 
        a python object (e.g. a dictionary).  This class does not
        create the deaddrop file, so you need to create a 'dummy' and then disribute
        the appropriate box id"""

    def __init__(self,config_file='config.json',box_file_id=DEADDROP_ID):
        self.box_file_id = box_file_id
        self.config_file = config_file
        self.box_config = JWTAuth.from_settings_file(self.config_file)
        self.box_client = Client(self.box_config)

    def _get_file_contents(self):
        result = self.box_client.file(self.box_file_id).content()
        return(result)        

    def _data_to_json_file(self,data,data_filename):
        with open(data_filename, 'w') as outfile:
            json.dump(data,outfile)

    def update_deaddrop_data_from_data(self,data,data_filename):
        self._data_to_json_file(data,data_filename)
        self.update_deaddrop_data_from_file(data_filename)

    def update_deaddrop_data_from_file(self,data_filename):
        deaddrop_file = self.box_client.file(self.box_file_id)
        updated_file_data = deaddrop_file.update_contents(data_filename)
        print(updated_file_data.get())                

    def get_deaddrop_data(self):
        """ Reads the dead drop data and returns a Python Object with the Data """
        return(json.loads(self._get_file_contents()))
 
def read_data_test(module_name):
    
    dead_drop = BoxDeadDrop()
    magnet_module_dictionary = dead_drop.get_deaddrop_data()
    
    return magnet_module_dictionary[module_name]
               
if __name__ == "__main__":
    read_data_test()
