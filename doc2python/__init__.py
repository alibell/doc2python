#
# doc2python script
# Extract text from doc file
#

# Dependencies
import os
import json
import compoundfiles
import re
from collections import OrderedDict

# For relative path management
package_directory = os.path.dirname(os.path.abspath(__file__))

# Functions

def bits_from_bytes(bytes_data, collapse = True):
    
    '''
        bits_from_bytes : get bits value from bytes
        Param :
            bytes_data : bytes value
            collapse : if false, the output is an array with the bits of each byte
    '''
    
    bits_data = []
    
    if type(bytes_data) == type(int()):
        bytes_data = [bytes_data]
        
    for byte in bytes_data:
        bit = bin(byte)[2:].rjust(8, '0')[::-1]
        bits_data.append(bit)
        
    if collapse:
        bits_data = "".join(bits_data)
        
    return(bits_data)

class bytesParser():
    '''
        bytesParser
            Parse byte string
            Get hexadecimal and decimal value for each
    '''
    def __init__(self):
        # Getting parser offsets from json files
        self._offset_names = ["FibBase","clw", "FibRgW97","cslw","FibRgFcLcb","cbRgFcLcb","FibRgLw97","cswNew", "FibRgCswNew", "pcd", "fc"]
        
        self._offsets_dict = {}

        for offset in self._offset_names:
            with open(package_directory+"/data/"+offset+".json", "r") as file:
                json_data = json.loads(file.read(), object_pairs_hook=OrderedDict)
                self._offsets_dict[offset] = json_data
                
    def _process_byte (self, byte):
        
        '''
            Get decimal and hexadecimal from a byte.
            Input :
                byte : byte value
            Output :
                dict containing byte, decimal and hexadecimal value
        '''
        
        processed_byte = {}
        processed_byte["bytes"] = byte
        processed_byte["decimal"] = int.from_bytes(byte, 'little')
        processed_byte["hexadecimal"] = hex(processed_byte["decimal"])
        
        return(processed_byte)
            
    def parse(self, data, data_type):
        '''
            Parse string
            input :
                - data, str : binary string to parse
                - data_type, str : name of the binary string (ex : Fib...)
            output :
                - dict of parsed data 
        '''
        
        if (data_type in self._offset_names):
            offset_table = self._offsets_dict[data_type]
            data_values = {} # Contains extracted value
            
            cursor = 0
            for offset in offset_table.keys():
                delta = int(offset_table[offset][1])
                
                if(len(offset_table[offset]) == 2):
                    data_values[offset] = self._process_byte(data[cursor:cursor+delta])
                    data_values[offset]["len"] = delta
                elif (offset_table[offset][2] == 'bits'):
                    bits_data = bits_from_bytes(data[cursor:cursor+delta])
                    var_name = offset_table[offset][3]
                    var_size = offset_table[offset][4]
                    
                    bits_dict = {}
                    cursor_bits = 0
                    for i in range(len(var_name)):
                        bits_dict[var_name[i]] = {}
                        bits_dict[var_name[i]]["bit"] = bits_data[cursor_bits:cursor_bits+var_size[i]][::-1]
                        bits_dict[var_name[i]]["numeric"] = int(bits_dict[var_name[i]]["bit"], 2)
                        bits_dict[var_name[i]]["length"] = var_size[i]
                        
                        cursor_bits += var_size[i]
                    
                    data_values.update(
                        bits_dict
                    )
                else:
                    raise Exception("Unknown data type")
                
                cursor += delta
                
            return(data_values)
        else:
            raise Exception("Unknown data type")
            
    def parseFib(self, wd_data):
        '''
            Parse Fib : header of Word Document
            input :
                - data, str : WordDocument data
            output :
                - dict of parsed data 
        '''
        
        fib_data = {} # Dict for output data
        
        cursor = 0
        # If 0 : need special rule
        steps = OrderedDict({
            "FibBase":32,
            "clw":2,
            "FibRgW97":28,
            "cslw":2,
            "FibRgLw97":88,
            "cbRgFcLcb":2,
            "FibRgFcLcb":0,
            "cswNew":2,
            "FibRgCswNew":0
        })
        
        for step in steps.keys():
            if (step == 'FibRgFcLcb'):
                delta = fib_data["cbRgFcLcb"]["cbRgFcLcb"]["decimal"]*8
            elif (step == 'FibRgCswNew'):
                delta = fib_data["cswNew"]["cswNew"]["decimal"]*2
            else:
                delta = steps[step]
            
            data = wd_data[cursor:cursor+delta]
            fib_data[step] = self.parse(data, step)

            cursor += delta # Move cursor
        
        # Post treatment
        
        ## Reconstitute FibRgCswNew
        if(fib_data["FibRgCswNew"]["nFibNew"]["decimal"] == 274):
            fib_data["FibRgCswNew"]["rgCswNewData"] = self._process_byte(fib_data["FibRgCswNew"]["nFibNew"]["bytes"]+fib_data["FibRgCswNew"]["rgCswNewData_extend"]["bytes"])
            fib_data["FibRgCswNew"]["rgCswNewData"]["len"] = 8         
            
        return fib_data
    
    def parsePlcPcd (self, table_PlcPcd_binary):
        '''
            Parse FlcPcd : PlcPcd from Table data
            input :
                - data, str : Table data
            output :
                - dict of parsed data 
        '''
        
        #Calculating number of CPs
        nb_cp = (len(table_PlcPcd_binary)+8)/12
        nb_apcd = nb_cp-1
        
        plcPcd_data = {
            "cp":[],
            "apcd":[]
        }
        
        cursor = 0
        
        # Parsing CPs
        for i in range(int(nb_cp)):
            delta = 4
            plcPcd_data["cp"].append(
                self._process_byte(table_PlcPcd_binary[cursor:cursor+delta])
            )
            
            cursor = cursor+delta
            
        #Parsing aPCDs
        for i in range(int(nb_apcd)):
            delta = 8
            plcPcd_data["apcd"].append(
                self.parse(table_PlcPcd_binary[cursor:cursor+delta], "pcd")
            )
            
            # Parsing fc
            plcPcd_data["apcd"][-1]["fc"] = self.parse(plcPcd_data["apcd"][-1]["fc"]["bytes"], "fc")
            
            cursor = cursor+delta
            
        return(plcPcd_data)
            
    def parsePcdt (self, table_pcdt_binary):
        '''
            Parse Pcdt : PlcPcd from pcdt data
            input :
                - data, str : Pcdt data
            output :
                - dict of parsed data 
        '''
        
        pcdt_data = {}
        
        steps = {
            "clxt":1,
            "lcb":4,
            "PlcPcd":0
        }
        cursor = 0
        
        for step in steps:
            if step == "PlcPcd":
                delta = len(table_pcdt_binary)-cursor
                pcdt_data[step] = {}
                pcdt_data[step] = self.parsePlcPcd(table_pcdt_binary[cursor:cursor+delta])
            else:
                delta = steps[step]
                pcdt_data[step] = self._process_byte(table_pcdt_binary[cursor:cursor+delta])
                
            pcdt_data[step]["len"] = delta
            cursor = cursor+delta
        
        return(pcdt_data)
    
    def parseClx (self, table_clx_binary):
        '''
            Parse Clx
            input :
                - data, str : Clx data
            output :
                - dict of parsed data 
        '''
        
        clx_data = {}
        
        # Getting RgPrc et Pcdt : Pcdt start by 0x02 while RgPrc doesn't, so whe search for 0x02
        for i in range(len(table_clx_binary)):
            if table_clx_binary[i] == 2:
                clx_data["RgPrc"] = table_clx_binary[0:i]
                clx_data["Pcdt"] = table_clx_binary[i:]
                
                break
                
        # Getting Pcdt data
        clx_data["Pcdt"] = self.parsePcdt(clx_data["Pcdt"])
        
        return(clx_data)
    
def process(file, encoding = "latin1"):
    '''
        Extract text from doc file.
        Input :
            file, str or bytesIO : file path ou bytesIO object for doc file
        Output :
            String of doc content
    '''
    
    # Loading parser
    bp = bytesParser()
    
    # Loading compound file data
    cf_data = compoundfiles.CompoundFileReader(file)
    
    # Getting data from cf
    WordDocument = cf_data.open("WordDocument").read()
    Table = cf_data.open([x.name for x in cf_data.root if x.name[1:]=='Table'][0]).read() # Fib lies sometimes on Table, so directly get it from cf_data

    # Parsing data
    FibData = bp.parseFib(WordDocument)
    
    # Getting Clx
    Clx = Table[FibData["FibRgFcLcb"]["fcClx"]["decimal"]:FibData["FibRgFcLcb"]["fcClx"]["decimal"]+FibData["FibRgFcLcb"]["lcbClx"]["decimal"]]
    Clx_parsed = bp.parseClx(Clx)
    
    # Getting Cps and PCDs
    cp = Clx_parsed["Pcdt"]["PlcPcd"]["cp"]
    apcd = Clx_parsed["Pcdt"]["PlcPcd"]["apcd"]
    
    # Getting text from doc
    text = []
    for i in range(len(apcd)):
        text_dict = {}
        text_dict["compressed"] = apcd[i]["fc"]['fCompressed']["numeric"]
        text_dict["start"] = int(apcd[i]["fc"]["fc"]["numeric"]/(1+text_dict["compressed"]))
        text_dict["end"] = int(text_dict["start"]+(2-text_dict["compressed"])*(cp[i+1]["decimal"]-cp[i]["decimal"]-1))
        text_dict["content"] = WordDocument[text_dict["start"]:text_dict["end"]]
        
        text.append(text_dict)
    
    # Return simple text array
    text_array = [x["content"].decode(encoding, errors = "ignore") for x in text]
    fulltext = "".join(text_array)
    
    # Postprocessing
    fulltext = re.sub("\\x13","",fulltext)
    fulltext = re.sub("\r","\r\n",fulltext)
    fulltext = re.sub('HYP?ERLINK "(.*?)"(?: *\\\\t *".*?")?(?: *\\\\o *".*?")?(?: *\\\\n *".*?")?(?: *\\\\m *".*?")?(?: *\\\\l *".*?")? *\\x14(.*?)\\x15','(\\2) [\\1]',fulltext)
    fulltext = re.sub("\\x00|\\x01|\\x14|\\x15","",fulltext)
    fulltext = re.sub('HYP?ERLINK *"(.*?)"','[\\1]',fulltext)
    fulltext = re.sub('INCLUDEPICTURE *"(.*?)"',"IMG[\\1]",fulltext)
    fulltext = re.sub('\\\\\* *MERGEFORMA(TINET?)','',fulltext)
    fulltext = re.sub("\\x07\\x07","\r\n", fulltext)
    fulltext = re.sub('\\x07','|',fulltext)
    
    return(fulltext)