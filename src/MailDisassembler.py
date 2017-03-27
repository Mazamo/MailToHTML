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
    
    The eml related code in this file is uses the default Email module.
"""

__author__ = "Nick de Visser"
__date__ = "2017-03-27"
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
import json
from __future__ import generators

def _windowsUnicode(string):
    if string is None:
        return None
    if sys.version_info[0] >= 3:  # Python 3
        return str(string, 'utf_16_le')
    else:  # Python 2
        return unicode(string, 'utf_16_le')

class Message(Object):
    def getData():
        raise NotImplementedError()

    def _assembleJSON():
        raise NotImplementedError()
    
    def factory(type):
        # Perform the test to determine wich Message file should be build.
        if type == "EMLMessage": return EMLMessage()
        if type == "MSGMessage": return MSGMessage()
        assert 0, "Bad Message creation: " + type
    factory = staticmethod(factory)


class EMLMessage(Message):
    def getData(self):
        pass

    def _assembleJSON(self):
        pass

    def _pullout(self, m, key):
        """
        Extracts content from an e-mail message.
        This works for multipart and nested multipart messages too.
        m   -- email.Message() or mailbox.Message()
        key -- Initial message ID (some String)
        Returns tuple(Text, Html, Files, Parts)
        Text  -- All text from all parts.
        Html  -- All HTMLS from all parts.
        Files -- Dictionary mapping extracted file to message ID it belongs to.
        Parts -- Number of parts in original message.
        """
        pass

    def _extract(self, messageFile, key):
        """
        Extracts all data from e-mail, includeing From, To, etc., and returns
        it as a dictionary.
        msgfile -- A file-like readable object
        key     -- Some ID string for that particular Message. Can be a file
                   name or anything.
        Returns a dict()
        Keys: from, to, subject, date, text, html, parts[, files]
        Key files will be pressent only when message contained binary files.
        For more see __doc__ for pullout() and caption() functions.
        """
        pass

    def _caption(self, origin)
        """
        Extracts: To, From, Subject and Date from email.Message() or 
        mailbox.Message()
        origin -- Message() object
        Returns tuple(From, To, Subject, Date)
        If message doesn't contain one/more of them, the empty strings will
        be returned.
        """
        pass

class MSGMessage(Message, OleFile.OleFileIO):
    def __init__(self, filename)
        OleFile.OleFileIO.__init__(self, filename)

    def _getStream(self, filename):
        """
        Check if there is a media stream.
        Every field in an OleMSG file is build op from text streams
        that must be handled seperatly.
        """
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

    @property
    def subject(self):
        return self._getStringStream('__substg1.0_0037')

    @property
    def _header(self):
        try:
            return self._header
        except Exeption:
            headerText = self._getStringStream('__substg1.0_007D')
            if headerText is not None:
                self._header = EmailParser().parsestr(headerText)
            else:
                self._header = None
            return selft._header

    @property
    def _date(self):
        # Get the message's header and extract the date
        if self.header is None:
            return None
        else:
            return self.header['date']

    @property
    def _parseDate(self):
        return email.utils.parsedate(self.date)

    @property
    def _sender(self):
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
                    result = result + "<" + email + ">"

            self._sender = result
            return result

    @property
    def _to(self):
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
    def _cc(self):
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
            self._cc = dislay
            return display

    @property
    def _body(self):
        # Get the message body
        return self_getStringStream('__substg1.0_1000')

    @property
    def _attachments(self):
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

            return self.attachments
    
    def _assembleJSON(self):
        attachmentNames = []
        # Save the attachments
        for attachment in self.attachments:
            attachmentNames.append(attachment.save())

        def xstr(s):
            return '' if s is None else str(s)

        emailObject = { 'from'          : xstr(self._sender),
                        'to'            : xstr(self._to),
                        'cc'            : xstr(self._cc),
                        'subject'       : xstr(self._subject),
                        'date'          : xstr(self._date),
                        'attachments'   : attachmentNames,
                        'body'          : decode_utf7(self.body)}
        
        return json.dumps(emailObject)
