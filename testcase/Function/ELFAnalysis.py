import argparse
import json
import logging
import os
from collections import defaultdict
from typing import Optional
# from elftools.common.py3combat import (
#     ifilter, byte2int, bytes2str, itervalues, str2bytes)
# from elftools.dwarf.die import DIE
# from elftools.elf.elffile import ELFFile

from .elf_wrapper import  ElfAddrObj


class ELFAnalysis:
    structDict = {}
    elfWrapper = None

    def loadPath(self, path):
        self.elfWrapper = ElfAddrObj(path)
        elffile = self.elfWrapper.elffile

        # dwarfinfo = elffile.get_dwarf_info()
        # self.parse_top_die_by_cu(dwarfinfo)

        if "_debugFlgStruct" in self.elfWrapper.struct_dict.keys():
            values = self.elfWrapper.struct_dict["_debugFlgStruct"]
            for key in values:
                self.structDict[key] = hex(
                    self.elfWrapper.get_var_addrs("debugFlgStruct." + key))

        for section in elffile.iter_sections():
            name = section.name
            if name == ".symtab":
                for cnt, symbol in enumerate(section.iter_symbols()):
                    # if symbol.name == 'CCU_PA01_AC_U':
                    #     print('a')
                    if symbol['st_info']['type'] == "STT_OBJECT":
                        self.structDict[symbol.name] = hex(symbol['st_value'])

    def getAddressWithName(self, name):
        return "0x" + hex(self.elfWrapper.get_var_addrs(name)).upper().replace("0X", "").zfill(8)
        # if self.elfWrapper is None:
        #     logging.error("ELF Wrapper is not initialized. Did you forget to call loadPath()?")
        #     return None  # 或者你也可以返回一个默认地址："0x00000000"
        # try:
        #     return "0x" + hex(self.elfWrapper.get_var_addrs(name)).upper().replace("0X", "").zfill(8)
        # except KeyError:
        #     logging.error(f"Variable not found in ELF file: {name}")
        #     return "0x00000000"

    def die_info_rec(self, die, name='', is_struct=False):

        if die.tag == 'DW_TAG_member' and die.attributes.get("DW_AT_name") and is_struct:
            member_name = die.attributes.get("DW_AT_name").value.decode()
            struct_val_name = name + "." + member_name
            try:
                self.structDict[struct_val_name.replace("debugFlgStruct.", "")] = hex(self.elfWrapper.get_var_addrs(struct_val_name))
            except:
                print(struct_val_name)

        if die.tag == 'DW_TAG_structure_type' and die.attributes.get("DW_AT_name"):
            name = die.attributes.get("DW_AT_name").value.decode().replace("_debugFlgStruct", "debugFlgStruct")
            if die.attributes.get("DW_AT_declaration") and die.attributes.get("DW_AT_declaration").value == 1:
                # print("struct {}: just a declaration".format(name))
                return
            for child in die.iter_children():
                self.die_info_rec(child, name, True)


    def parse_top_die_by_cu(self, dwarfinfo):
        for CU in dwarfinfo.iter_CUs():
            top_DIE = CU.get_top_DIE()
            if "DEBUG_" in top_DIE.attributes.get("DW_AT_name").value.decode().upper():
                for child in top_DIE.iter_children():
                    try:
                        self.die_info_rec(child, False)
                    except:
                        pass
#
if __name__ == '__main__':
    elfAnalysis = ELFAnalysis()
    elfAnalysis.loadPath(r"CYT4BF_M7_Master.out")
    # add1 = elfAnalysis.getAddressWithName('CCU_PA01_AC_U.SI_u8_AC_MHU_BlowerSpd')
    print(elfAnalysis.structDict)