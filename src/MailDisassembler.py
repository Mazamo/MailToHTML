#!/usr/bin/env python
#-*- coding: latin-1 -*-

"""
MailDisassembler:
    Extracts the data from MSG and EML files.
    This class can be used in conjunctiong with the HTMLAssembler to 
    create a basic webpage that can display the contents of a email message.

    The msg related code in this file is based on the msg-extractor 
    from Matthew Walker:
    https://github.com/mattgwwalker/msg-extractor/blob/master/ExtractMsg.py   

    Github Page:
    https://github.com/Mazamo/MailToHTML 
"""

__author__ = "Nick de Visser"
__date__ = "2017-04-10"
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
import email
import re
import olefile as OleFile
import json
from email.parser import Parser as EmailParser
from emaildata.metadata import MetaData
from emaildata.text import Text
from emaildata.attachment import Attachment

class BaseMessage(object):
    """
    Abstract class that serves as a base framework of an message.
    """
    
    EXTENSION = None

    def __init__(self, filename):
        self._filename = filename

    def getData(self):
        raise NotImplementedError()

    def _assembleJSON(self):
        def xstr(s):
            return '' if s is None else str(s)

        emailObject = { 'from'          : xstr(self.sender),
                        'to'            : xstr(self.to),
                        'cc'            : xstr(self.cc),
                        'subject'       : xstr(self.subject),
                        'date'          : xstr(self.date),
                        'attachments'   : self.attachments,
                        'body'          : self.body}
        
        return json.dumps(emailObject)
    
    @property
    def sender(self):
        raise NotImplementedError()

    @property
    def to(self):
        raise NotImplementedError()

    @property
    def cc(self):
        raise NotImplementedError()

    @property
    def subject(self):
        raise NotImplementedError()

    @property
    def date(self):
        raise NotImplementedError()

    @property
    def body(self):
        raise NotImplementedError()

    @property
    def attachments(self):
        raise NotImplementedError()

    @classmethod
    def check_extension(cls, extension):
        return extension == cls.EXTENSION

class EMLMessage(BaseMessage):
    """
    EML message extractor that extracts the following data form a eml file:
    - Source name
    - Source email address
    - Destinatoin email address
    - CC
    - Date
    - Content
    - A list of attachment names
    """
	
    EXTENSION = ".eml"

    def init(self, filename):
        super(EMLMessage, self).__init__()

        self._text = ""
        self._data = None
        self._message = None
	
    def getData(self):
        f = open(self._filename)
        self._message = email.message_from_file(f)
        f.close()

        self._data = MetaData(self._message).to_dict()

        return self._assembleJSON()

    @property
    def body(self):
        if Text.text(self._message) is not None:
            return Text.text(self._message)
        else:
            return ""

    @property
    def attachments(self):
    	attachmentList = []
    	for content, filename, mimetype, message in Attachment.extract(self._message):
    	    attachmentList.append(filename)
    	return attachmentList

    @property
    def sender(self):
    	if self._data.get('sender') is not None:
    	    return self._data.get('sender')

    @property
    def to(self):
    	receiverList = []
    	for receiver in self._data.get('receivers'):
        	if receiver is not None:
        	    receiverList.append(receiver)
    	return receiverList

    @property
    def cc(self):
    	ccList = []
    	for c in self._data.get('cc'):
    	    if c is not None:
    	        ccList.append(c)
    	return ccList

    @property
    def subject(self):
    	if self._data.get('subject') is not None:
    	    return self._data.get('subject')

    @property
    def date(self):
    	if self._data.get('date') is not None:
    	   tempDate = (self._data.get('date')).timetuple()
    	   date = "{0}-{1}-{2}, {3}:{4}".format(tempDate[2], tempDate[1], tempDate[0], tempDate[3], tempDate[4])
    	   return date

class MSGMessage(BaseMessage, OleFile.OleFileIO):
    """
    MSG message extractor that gets the following data form a msg file:
    - Source name
    - Source email address
    - Destinatoin email address
    - CC
    - Date
    - Content
    - A list of attachment names
    """
    
    EXTENSION = '.msg'

    def __init__(self, filename):
        OleFile.OleFileIO.__init__(self, filename)

    def _getStream(self, filename):
        """
        Check if there is a media stream.
        Every field in an OleMSG file is build op from text streams
        that must be handled seperatly.
        """
        if self.exists(filename):
            stream = self.openstream(filename)
            return stream.read()
        else:
            return None

    def _windowsUnicode(self, string):
        if string is None:
            return None
        if sys.version_info[0] >= 3:  # Python 3
            return str(string, 'utf_16_le')
        else:  # Python 2
            return unicode(string, 'utf_16_le')

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
        unicodeVersion = self._windowsUnicode(self._getStream(filename + '001F'))
        if asciiVersion is None:
            return unicodeVersion
        elif unicodeVersion is None:
            return asciiVersion
        else:
            if prefer == 'unicode':
                return unicodeVersion
            else:
                return asciiVersion

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def header(self):
        try:
            return self._header
        except Exception:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return self._header

    @property
    def date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def _parseDate(self):
        return email.utils.parsedate(self.date)

    @property
    def sender(self):
        try:
            return self._sender
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["from"]
                if headerResult is not None:
                    self._sender = headerResult
                    return headerResult

            # Extract from other fields
            text = self._getStringStream('__substg1.0_0C1A')
            email = self._getStringStream('__substg1.0_0C1F')
            result = None

            if text is None:
                result = email
            else:
                result = text
                if email is not None:
                    result = result + email

            self._sender = result
            return result

    @property
    def to(self):
        try:
            return self._to
        except Exception:
            # Chck header first
            if self.header is not None:
                headerResult = self.header["to"]
                if headerResult is not None:
                    self._to = headerResult
                    return headerResult

            # Extract from other fields
            display = self._getStringStream('__substg1.0_0E04')
            self._to = display
            return display


    @property
    def cc(self):
        try:
            return self._cc
        except Exception:
            # Check header first
            if self.header is not None:
                headerResult = self.header["cc"]
                if headerResult is not None:
                    self._cc = headerResult
                    return headerResult

            # Extract from other fields
            display = self._getStringStream('__substg1.0_0E03')
            self._cc = display
            return display

    @property
    def body(self):
        # Get the message body
        return self._getStringStream('__substg1.0_1000') 

    @property
    def attachments(self):
        class Attachment(object):
            def __init__(self, msg, dir_):
                self.longFilename = msg._getStringStream([dir_, '__substg1.0_3707'])
                self.shortFilename = msg._getStringStream([dir_, '__substg1.0_3704'])
                self.data = msg._getStream([dir_, '__substg1.0_37010102'])

        try:
            return self._attachments
        except Exception:
            # Get the attachments
            attachmentDirs = []

            for dir_ in self.listdir():
                if dir_[0].startswith('__attach') and dir_[0] not in attachmentDirs:
                    attachmentDirs.append(dir_[0])

            self._attachments = []
            for attachmentDir in attachmentDirs:
                self._attachments.append(Attachment(self, attachmentDir))

            attachmentNames = []
            for attachment in self._attachments:
                attachmentNames.append(attachment.longFilename)
            
            return attachmentNames
    
    def getData(self):
        return self._assembleJSON()

class MessageManager(object):
    """
    Manager class that manages the extraction from the message.
    This class chooses the right extractor class and returns the 
    assembled JSON file.
    """
    
    ARCHIVES_ENGINES = [EMLMessage, MSGMessage]

    def __init__(self, filename):
        self._filename, self._extension = os.path.splitext(filename)
        self._messageEngine = self.chooseMessageEngine()


    def chooseMessageEngine(self):
        for engine in self.ARCHIVES_ENGINES:
            if engine.check_extension(self._extension):
                return engine(self._filename + self._extension)

    def createJSON(self):
        return self._messageEngine.getData()

if __name__ == '__main__':
    message = MessageManager("test_msg.msg")
    print message.createJSON()

    print ""

    message = MessageManager("unicode.msg")
    print message.createJSON()

    print "" 

    message = MessageManager("test_sample_message.eml")
    print message.createJSON()

    print ""

    message = MessageManager("message.eml")
    print message.createJSON()

    print ""

    message = MessageManager("email_test.msg")
    print message.createJSON()
