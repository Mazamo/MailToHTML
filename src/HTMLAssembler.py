#!/usr/bin/env python
#-*- coding: latin-1 -*-
"""
HTMLAssembler:
    Builds an html page and fills it with an given JSON file.
    The build HTML page is a hardcoded site that will be filled
    JSON file to represent a e-mail message.

    The site will be build in the following format:
    ======= MESSAGE BEGIN =======
    FROM	: 
    TO  	:
    CC  	:
    Subject	:
    Date	:

    CONTENT

    ATTACHMENT LIST:
    example_one.pdf
    example_two.tif
    ======= MESSAGE  END =======
"""

