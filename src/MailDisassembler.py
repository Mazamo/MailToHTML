#!/usr/bin/env python
#-*- coding: latin-1 -*-
"""
MailDisassembler:
    Extracts the data from MSG and EML files.
    This class can be used in conjunctiong with the HTMLAssembler to 
    create a basic webpage that can display the contents of a email message.

    The msg related code in this class is based on the msg-extractor 
    from Matthew Walker:
    https://github.com/mattgwwalker/msg-extractor/blob/master/ExtractMsg.py   
    
    Github Page:
    https://github.com/Mazamo/MailToHTML 
"""

__author__ = "Nick de Visser"
__date__ = "2017-03-17"
__version__ = '0.0'

# --- LICENSE -----------------------------------------------------------------
#
#    Copyright 2017 Nick de Visser
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.

import os
import sys
import glob
import traceback
from email.parser import Parser as EmailParser
import email.utils
import email.utils
import olefile as OleFile

def windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')

class Message():

class EMLMessage(Message):


class MSGMessage(Message, OleFile.OleFileIO):
    def __init__(self, filename)

    def _getStream(self, filename):
        if self.exist(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _getStringStream(self, filename, prefer='unicode'):
        """Gets a string representation of the requested filename.
        Checks for both ASCII and Unicode representations and returns
        a value if possible. If there are both ASCII and Unicode
        versions, then the parameter /prefer/ specifies which will be
        returned.
        """

        if isinstance(filename, list):
            # Join with slashes to make it easier to append the type
            filename = "/".join(filename)

        asciiVersion = self._getStream(filename + '001E')
        unicodeVersion = windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion


