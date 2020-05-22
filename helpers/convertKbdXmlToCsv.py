#-*- coding: utf-8 -*-

""" convertKbdXmlToCsv.py
	author:  Lucas Becker
	version: 2019-07-05
	Python version: 3.x (mainly because the csv module does not support Unicode in 2.x)

	Convert the keyboard shortcut XML to CSV so we can open it in Excel
	for better viewing.
"""

import xml.etree.ElementTree as ET
import csv
import os, sys

# put the current scheme version GS is using in here
# you find it in the 'xmlns' attribute at root level
SCHEME = "{http://www.graphisoft.com/Development/WorkEnvironment/Scheme/2018/01}"
kbd_folder_name = 'Tastaturkürzel'  # GER name of exported folder

dirpath = os.path.dirname(__file__)  # location of this script
# inputs are not working with Sublime Text so I just hardcode the target :/
# target =  str(input("Which workspace keyboard shortcuts should be converted? "))
target = "23"
source = dirpath.split("\\")[:-1]
source.extend(["workspaces", target, kbd_folder_name])
source = ("\\").join(source)

# get the XML path
files = []
#! r=root, d=directory, f=files
for r, d, f in os.walk(source):
	for file in f:
		if '.xml' in file:
			files.append(os.path.join(r, file))

if len(files) <= 0:
	print("There is no keyboard shorcuts XML.")
else:
	file = files[0]

# parse the XML
tree = ET.parse(file)
root = tree.getroot()

csv_filepath = source + "\\keyboard-shortcuts.csv"
# open a file for writing
try:
	# newline='' important in py3 so we get no empty lines
	# uft-8 is not understood by Excel... sigh
	kbdshortcuts = open(csv_filepath, 'w', newline='', encoding='Windows-1252')  
except:
	print("Failed to open the output file.")
	raise

# create the csv writer object
csvwriter = csv.writer(kbdshortcuts, dialect='excel')

for child in root.findall(".//{}CommandDescriptor[@comment]".format(SCHEME)):
	shortcutrow = []

	###  v   get command name
	command_name = child.attrib["comment"].replace("&", "")  # get rid of '&'
	command_name = command_name.replace("  ", " & ")  # put the '&' right back in when actually needed
	shortcutrow.append(command_name)

	###  v   get keyboard shortcuts
	# modifiers.split() > character > specialKey
	shortcut = child.find(".//{}Shortcut".format(SCHEME))
	try:
		char = shortcut.attrib["character"]
	except:
		char = ""

	try:
		mod = shortcut.attrib["modifiers"].split("|")
	except:
		mod = ""

	try:
		special = shortcut.attrib["specialKey"]
	except:
		special = ""

	keys = (" + ").join(filter(None, [*mod, char, special]))
	shortcutrow.append(keys)

	# write row
	csvwriter.writerow(shortcutrow)
kbdshortcuts.close()
