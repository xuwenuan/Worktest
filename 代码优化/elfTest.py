import argparse
import json
import os
from collections import defaultdict
from typing import Optional
from elftools.common.py3compat import (
    ifilter, byte2int, bytes2str, itervalues, str2bytes)
from elftools.dwarf.die import DIE
from elftools.elf.elffile import ELFFile

from elf_wrapper import ElfAddrObj

Map_TypePrefix = {
    'DW_TAG_base_type': '',
    'DW_TAG_structure_type': 'struct ',
    'DW_TAG_union_type': 'union ',
    'DW_TAG_pointer_type': 'pointer '
}

Map_AnonTypes = {
    'DW_TAG_subroutine_type': 'subroutine',
    'DW_TAG_pointer_type': 'pointer',
    'DW_TAG_union_type': 'union'
}


# # recursive function to get type of a DIE node
# def die_type_rec(die: DIE, prev: Optional[DIE]):
#     t = die.attributes.get("DW_AT_type")
#     print(t)
#
#     if t.form == 'DW_FORM_ref4':
#         ref = t.value
#         ref_die = dwarfinfo.get_DIE_from_refaddr(ref + die.cu.cu_offset)
#         #ref_die = dwarfinfo.get_DIE_from_attribute(ref)
#         return ref + die.cu.cu_offset
#     # else:
#     #     # print(die)
#     #     prefix = '* ' if prev.tag == 'DW_TAG_pointer_type' else ''
#     #
#     #     # got a type
#     #     if die.attributes.get("DW_AT_name"):
#     #         # common named type with prefix
#     #         return prefix + Map_TypePrefix.get(die.tag, f'unknown: {die.tag}') \
#     #                + die.attributes.get("DW_AT_name").value.decode()
#     #     elif die.tag == 'DW_TAG_structure_type' and prev.tag == 'DW_TAG_typedef':
#     #         # typedef-ed anonymous struct
#     #         return prefix + 'struct ' + prev.attributes.get("DW_AT_name").value.decode()
#     #     else:
#     #         # no name types
#     #         return prefix + Map_AnonTypes.get(die.tag, f'unknown: {die.tag}')

# recursive function to get all struct members
def die_info_rec(die: DIE, name='', is_struct=False, structName = "", structAddress=0x0):
    # print(die)
    if die.tag == 'DW_TAG_member' and die.attributes.get("DW_AT_name") and is_struct:
        member_name = die.attributes.get("DW_AT_name").value.decode()
        #member_type = die_type_rec(die, None)
        offset = 0x0
        if die.attributes.get("DW_AT_data_member_location"):
            member_offset = die.attributes.get("DW_AT_data_member_location").value
            print('  > .{}, offset: {}'.format(member_name, member_offset))
        if die.attributes.get("DW_AT_bit_size") and die.attributes.get("DW_AT_bit_offset"):
            member_bit_size = die.attributes.get("DW_AT_bit_size").value
            member_bit_offset = die.attributes.get("DW_AT_bit_offset").value
            print('  > .{}, , bit_offset: {}, bit_size: {}'.format(member_name,  member_bit_size, member_bit_offset))

        #save to return data
        # if member_type.startswith('*'):
        #     # pointer member, change to *name -> type
        #     struct_data[name]['*' + member_name] = member_type[1:]
        # else:
        struct_data[name + "." + member_name] = member_bit_offset

    if die.tag == 'DW_TAG_structure_type' and die.attributes.get("DW_AT_name"):
        name = 'struct ' + die.attributes.get("DW_AT_name").value.decode()
        if die.attributes.get("DW_AT_declaration") and die.attributes.get("DW_AT_declaration").value == 1:
            print("struct {}: just a declaration".format(name))
            return

        size = die.attributes.get("DW_AT_byte_size").value
        # print("{}, size:{}".format(name, size))

        # recursion into all children DIE
        for child in die.iter_children():
            die_info_rec(child, name, True, "", 0x0)


def parse_top_die_by_cu(dwarfinfo):
    j = 0
    for CU in dwarfinfo.iter_CUs():
        j = j + 1
        #print('  Found a compile unit at offset %s, length %s' % (CU.cu_offset, CU['unit_length']))

        # Start with the top DIE, the root for this CU's DIE tree
        top_DIE = CU.get_top_DIE()

        #print("------------------------Top Die[{}] start-----------------------------------------".format(j))

        if "DEBUG_" in top_DIE.attributes.get("DW_AT_name").value.decode().upper():
            # print(top_DIE)
            # Display DIEs recursively starting with top_DIE
            i = 0
            for child in top_DIE.iter_children():
                # for child in CU.iter_DIEs():
                i = i + 1
                #print("Top Die[{}]->child[{}]:", j, i)
                try:
                    die_info_rec(child, False, "", 0x0)
                except:
                    pass

        #print("------------------------Top Die[{}] end-----------------------------------------".format(j))


# dict for all struct members
struct_data = defaultdict(dict)

# elf_file = ".\\CYT4BF_M7_Master.out"
#
# print('Processing file:', elf_file)
# f = open(elf_file, 'rb')
# elffile = ELFFile(f)
#
# if not elffile.has_dwarf_info():
#     print(f'ERROR: input file {elf_file} has no DWARF info')
#     exit(1)
#
# dwarfinfo = elffile.get_dwarf_info()
#
# parse_top_die_by_cu(dwarfinfo)
#
# f.close()

elf = ElfAddrObj(r"CYT4BF_M7_Master.out")
var_addr = elf.get_var_addrs('debugFlgStruct.VbOUT_SSB_SSBOut_flg_old')
print(var_addr)