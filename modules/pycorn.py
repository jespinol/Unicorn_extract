#!/usr/bin/env python
# -*- coding: utf-8 -*-
'''
PyCORN - script to extract data from .res (results) files generated
by UNICORN Chromatography software supplied with ÄKTA Systems
(c)2014-2016 - Yasar L. Ahmed
v0.18
'''

from __future__ import print_function

import io
import struct
import xml.etree.ElementTree as ET
from collections import OrderedDict
from zipfile import ZipFile
from zipfile import is_zipfile


class pc_uni6(OrderedDict):
    '''
    A class for holding the pycorn/RESv6 data
    A subclass of `dict`, with the form `data_name`: `data`.
    '''
    # for manual zip-detection
    zip_magic_start = b'\x50\x4B\x03\x04\x2D\x00\x00\x00\x08'
    zip_magic_end = b'\x50\x4B\x05\x06\x00\x00\x00\x00'

    # hack to get pycorn-bin to move on
    SensData_id = 0
    SensData_id2 = 0
    Fractions_id = 0
    Fractions_id2 = 0

    def __init__(self, inp_file):
        OrderedDict.__init__(self)
        self.file_name = inp_file
        self.inject_vol = 0.0
        self.run_name = 'blank'

    def load(self, show=False):
        '''
        zip-files inside the zip-bundle are replaced by dicts, again with dicts with filename:content
        Chrom.#_#_True (=zip-files) files are unpacked from binary to floats by unpacker()
        To access x/y-value of Chrom.1_2:
        udata = pc_uni6("mybundle.zip")
        udata.load()
        x = udata['Chrom.1_2_True']['CoordinateData.Volumes']
        y = udata['Chrom.1_2_True']['CoordinateData.Amplitudes']
        '''
        with open(self.file_name, 'rb') as f:
            input_zip = ZipFile(f)
            zip_data = self.zip2dict(input_zip)
            self.update(zip_data)
            proc_yes = []
            proc_no = []
            for i in self.keys():
                tmp_raw = io.BytesIO(input_zip.read(i))
                f_header = tmp_raw.read(9)
                # tmp_raw.seek(0)
                # the following if block is to fix the non-standard zip files
                # by stripping out all the null-bytes at the end
                # see https://bugs.python.org/issue24621
                if f_header == self.zip_magic_start:
                    proper_zip = tmp_raw.getvalue()
                    f_end = proper_zip.rindex(self.zip_magic_end) + 22
                    tmp_raw = io.BytesIO(proper_zip[0:f_end])
                if is_zipfile(tmp_raw):
                    tmp_zip = ZipFile(tmp_raw)
                    x = {i: self.zip2dict(tmp_zip)}
                    self.update(x)
                    proc_yes.append(i)
                else:
                    pass
                    proc_no.append(i)
            if show:
                print("Loaded " + self.file_name + " into memory")
                print("\n-Supported-")
                for i in proc_yes:
                    print(" " + i)
                print("\n-Not supported-")
                for i in proc_no:
                    print(" " + i)
        # filter out data we dont deal with atm
        to_process = []
        for i in self.keys():
            if "Chrom" in i and not "Xml" in i:
                to_process.append(i)
        if show:
            print("\nFiles to process:")
            for i in to_process:
                print(" " + i)
        for i in to_process:
            for n in self[i].keys():
                if "DataType" in n:
                    a = self[i][n]
                    b = a.decode('utf-8')
                    x = b.strip("\r\n")
                else:
                    x = self.unpacker(self[i][n])
                tmp_dict = {n: x}
                self[i].update(tmp_dict)
        if show:
            print("Finished decoding x/y-data!")

    @staticmethod
    def zip2dict(inp):
        '''
        input = zip object
        outout = dict with filename:file-object pairs
        '''
        mydict = {}
        for i in inp.NameToInfo:
            tmp_dict = {i: inp.read(i)}
            mydict.update(tmp_dict)
        return (mydict)

    @staticmethod
    def unpacker(inp):
        '''
        input = data block
        output = list of values
        '''
        read_size = len(inp) - 48
        values = []
        for i in range(47, read_size, 4):
            x = struct.unpack("<f", inp[i:i + 4])
            x = x[0]
            values.append(x)
        return (values)

    def xml_parse(self, show=False):
        '''
        parses parts of the Chrom.1.Xml and creates a res3-like dict
        '''
        tree = ET.fromstring(self['Chrom.1.Xml'])
        mc = tree.find('Curves')
        me = tree.find('EventCurves')
        # print(tree.tag)
        # print(tree.attrib)
        event_dict = {}
        for i in range(len(me)):
            magic_id = self.SensData_id
            e_type = me[i].attrib['EventCurveType']
            e_name = me[i].find('Name').text
            if e_name == 'Fraction':
                e_name = 'Fractions'  # another hack for pycorn-bin
            e_orig = me[i].find('IsOriginalData').text
            e_list = me[i].find('Events')
            e_data = []
            for e in range(len(e_list)):
                e_vol = float(e_list[e].find('EventVolume').text)
                e_txt = e_list[e].find('EventText').text
                e_data.append((e_vol, e_txt))
            if e_orig == "false":
                pass
                # print("not added - not orig data")
            if e_orig == "true":
                # print("added - orig data")
                x = {'run_name': "Blank", 'data': e_data, 'data_name': e_name, 'magic_id': magic_id}
                event_dict.update({e_name: x})
        self.update(event_dict)
        chrom_dict = {}
        for i in range(len(mc)):
            d_type = mc[i].attrib['CurveDataType']
            d_name = mc[i].find('Name').text
            d_fname = mc[i].find('CurvePoints')[0][1].text
            d_unit = mc[i].find('AmplitudeUnit').text
            magic_id = self.SensData_id
            try:
                x_dat = self[d_fname]['CoordinateData.Volumes']
                y_dat = self[d_fname]['CoordinateData.Amplitudes']
                zdata = list(zip(x_dat, y_dat))
                if d_name == "UV cell path length":
                    d_name = "xUV cell path length"  # hack to prevent pycorn-bin from picking this up
                x = {'run_name': "Blank", 'data': zdata, 'unit': d_unit, 'data_name': d_name, 'data_type': d_type,
                     'magic_id': magic_id}
                chrom_dict.update({d_name: x})
            except:
                KeyError
                # don't deal with data that does not make sense atm
                # orig2.zip contains UV-blocks that are (edited) copies of
                # original UV-trace but they dont have the volume data
            if show:
                print("---")
                print(d_type)
                print(d_name)
                print(d_fname)
                print(d_unit)
        self.update(chrom_dict)

    def clean_up(self):
        '''
        deletes everything and just keeps relevant run-date
        resulting dict is more like res3
        '''
        manifest = ET.fromstring(self['Manifest.xml'])
        for i in range(len(manifest)):
            file_name = manifest[i][0].text
            self.pop(file_name)
        self.pop('Manifest.xml')
