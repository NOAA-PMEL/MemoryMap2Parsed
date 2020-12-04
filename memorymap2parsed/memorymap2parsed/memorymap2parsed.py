"""Main module."""
import pandas as pd
from pathlib import Path
import re

def convert_excel_mem(path):
    size_exp = r"\[([A-Za-z0-9_]+)\]"

    xls = pd.ExcelFile(path)
    ram_list = pd.DataFrame()

    for sheet in xls.sheet_names:
        print(sheet)
        if sheet.find('MEM') >= 0:
            ram = pd.read_excel(xls, sheet_name=sheet, usecols='C:S', header=None, skiprows=2, )
            if ram.shape[1] < 16:
                for i in range(ram.shape[1],17):
                    ram[i] = pd.NA
            
            
            row_num = 0
            for idx, row in ram.iterrows():
                if idx == 0:
                    row0 = row
                else:
                    for i in range(3, ram.shape[1]+2):
                        addr = int(row[2],16) + int(row0[i], 16)
                        
                        if not pd.isna(row[i]):
                            m = re.search(size_exp, row[i])
                            try:
                                length = int(int(m.group(1)[1:3])/8)
                            except:
                                length = -999
                            v_type = m.group(1)[0]
                            var_name = row[i][:m.start()-1]

                            var_name = var_name.replace(' ', '_')
                            rdict = {
                                    'name':var_name,
                                    'start_addr':int(addr),
                                    'num_bytes':int(length),
                                    'type':v_type
                                    }
                            ram_list = ram_list.append(rdict, ignore_index=True)
                            

    ram_list = ram_list.astype({'num_bytes': 'uint8'})
    ram_list = ram_list.astype({'start_addr':'uint32'})
    print(ram_list)





if __name__ == '__main__':
    path = Path('memorymap2parsed/tests/files/Example.xlsx')

    convert_excel_mem(path)


    # , nrows=16, usecols='C:S', header=None)