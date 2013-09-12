#!/usr/bin/env python3
# -*- coding: utf-8 -*-	
#
# This is a filter to convert Visual Basic v6.0 code
# into something doxygen can understand.
# Copyright (C) 2005  Basti Grembowietz
# 
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
# 
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# 
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
# ------------------------------------------------------------------------- 
#
# This filter depends on following circumstances:
# in VB-code,
#  '! comments get converted to doxygen-class-comments (comments to a class)
#  '* comments get converted to doxygen-comments (function, sub etc)
#
#
# v0.1 - 2004-12-25
#  initial work
# v0.2 - 2004-12-30
#  added states
# v0.3 - 2004-12-31
#  removed states =)
# v0.4 - 2005-01-01
#  added class-comments
# v0.5 - 2005-01-03
#  changed default behaviour from "private" to "public"
#  + fixed re_doxy (whitespace now does not matter anymore)
#  + fixed re_sub and re_func (brackets inside brackets ok now)
# v0.6 - 2005-02-14
#  minor changes
# v0.7 - 2005-02-23
#  refactoring: removed double code.
#  + VB-Types are enabled now
#  + Doxygen-Comments can also start in the line of the feature which should be documented
# v0.8 - 2005-02-25
#  changed command line switches: now the usage is just "vbfilter.py filename".
# v0.9 - 2005-03-09
#  added handling of friends in vb.
# v0.10 - 2005-04-14
#  added handling of Property Let and Set
#  added recognition of default-values for parameters
# v0.11 - 2005-05-05
#  fixed handling of Property Get ( instead of Set ... )
# ========================================================================= 
# 2008/2/26 modified by Ryo Satsuki
#  modified handling of variable (Const, initial value, array)
#  modified handling of Function for Variant-return-function
#  added handling of End Function/Sub
#  added handling of Enum
#  added handling of blank line to keep comment block separation
# 2008/2/28 modified by Ryo Satsuki
#  modified handling of Function / Sub so as to format args
#  added handling of multiple divided lines
# 2008/4/9 modified by Ryo Satsuki
#  modified handling of comment for "'" in strings
#  added handling of a double quotation marks in a strings
#  modified handling of initial values so as to pass expressions
# 2008/8/27 modified by Ryo Satsuki
#  corrected handling of property procedure
#  modified handling of Sub so as to handle Property Set procedure
#  modified handling of Enum
#==========================================================================
# 2011/03/16 modified by SuzumeJr
#	convert python3.2
#	suport right side doxygen-comment
#	suport block doxygen-comment
#	fixed Lost WithEvents member
#	fixed put member type
#	suport handling of Event
#	suport option Puts Form Controls
#	fiexed class-blockcomment
#2011/03/25
#	Trouble to which the function is not output is corrected when there is # in an initial value in an optional argument of the function. 
#	Trouble to which two class comments are output is corrected. 

import getopt          # get command-line options
import os.path         # getting extension from file
import string          # string manipulation
import sys             # output and stuff
import re              # for regular expressions

# VB source encoding (added by R.S.)
src_encoding = "cp932"

# regular expression
## re to strip comments from file (modified by R.S.)
re_comments   = re.compile(r"((?:\"(?:[^\"]|\"\")*\"|[^\"'])*)'.*")
re_VB_Obj    = re.compile(r"\s*BEGIN\s*([\w.]*)\s+(\w*)", re.I)
re_VB_Obj_St = re.compile(r"\s*BEGIN\s+([\w.]*)\s+(\w*)", re.I)
re_VB_Obj_Ed = re.compile(r"^\s*End$")
re_VB_Obj_Pr = re.compile(r"^\s*(\w*)\s*=\s*[^\s].*$")
re_VB_Name    = re.compile(r"\s*Attribute\s+VB_Name\s+=\s+\"(\w+)\"", re.I)
re_VB_Attrib  = re.compile(r"\s*Attribute", re.I)

## re to blank line (added by R.S.)
re_blank_line = re.compile(r"^\s*$")

## re to search doxygen-class-comments (modified by R.S.)
re_doxy_class = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*'!(.*)")
## re to search doxygen-block-comments
re_doxy_block_st = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*'/\*\*(.*)")
re_doxy_block_proc = re.compile(r"(.*)'(.*)")
re_doxy_block_ed = re.compile(r"(.*)\*/(.*)")
## re to search doxygen-comments (modified by R.S.)
re_doxy       = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*''(.*)")
re_doxy_bk    = re.compile(r"(?:\"(?:[^\"]|\"\")*\"|[^\"'])*'<(.*)")
## re to search for global variables members (used in bas-files)
re_globals    = re.compile(r"\s*Global\s+(Const\s+)?([^']+)", re.I)
## re to search for class-members (used in cls-files) (modified by R.S.)
re_members    = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(?:(Const\s+)?(?:WithEvents\s+)?(?:Dim\s+)?([\w]+(?:\([\w\s\(\)\+\-\*/\.]*\))?)\s+As\s+([\w.]+)\s*(?:=\s*(\"(?:[^\"]|\"\")*\"|[^']+))?|(?:Const\s+([\w\(\)]+)\s+=\s*(\"(?:[^\"]|\"\")*\"|[^']+)))", re.I)
re_event	  = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}Event\s+(\w+)\s*(\([\w\s=,\(\)\+\-\*/\.\"]*\))", re.I)
re_array      = re.compile(r"([\w]+)\(([\w\s\(\)\+\-\*/\.]*)\)", re.I)
re_const_string	= re.compile(r"\"(?:[^\"]|\"\")*\"")
re_backslash	= re.compile(r"\\")
re_doublequote	= re.compile(r"(?=.)\"\"(?=.)")
## re to search Propertys
re_property     = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(Property\s+(?:Get|Let|Set))\s+(\w+)\s*(\([\w\s=,\(\)\+\-\*/\.\"]*\))(?:\s+As\s+(\w+))?", re.I)
re_endProperty  = re.compile(r"End\s+(?:Property)", re.I)
## re to search Subs (modified by R.S.)
re_sub        = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(Sub)\s+(\w+)\s*(\([\w\s=,\(\)\+\-\*/\.\"]*\))", re.I)
re_endSub  	  = re.compile(r"End\s+(?:Sub)", re.I)
## re to search Functions (modified by R.S.)
re_function = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}(Function)\s+(\w+)\s*(\([\w\s=,#\(\)\+\-\*/\.\"]*\))(?:\s+As\s+(\w+))?", re.I)
re_endFunction = re.compile(r"End\s+(?:Function)", re.I)
## re to search args (added by R.S.)
re_arg      = re.compile(r"\s*(Optional\s+)?((?:ByVal\s+|ByRef\s+)?(?:ParamArray\s+)?)(\w+)(\(\s*\))?(?:\s+As\s+(\w+))?(?:\s*=\s*(\"(?:[^\"]|\"\")*\"|[^,\)]+))?", re.I)
## re to search for type-statements
re_type     = re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}Type\s+(\w+)", re.I)
## re to search for type-statements
re_endType  = re.compile(r"End\s+Type", re.I)
## re to search for enum  (added by R.S.)
re_enum		= re.compile(r"\s*(Public\s+|Friend\s+|Private\s+|Static\s+){0,1}Enum\s+(\w+)", re.I)
re_endEnum  = re.compile(r"End\s+Enum", re.I)

# default "level" (private / public / protected) to take when not specified
def_level = "public:"

# strips vb-style comments from string
def strip_comments(str):
	my_match = re_comments.match(str)
	if my_match is not None:
		return my_match.group(1)
	else:
		return str

def doxy_back_comments(str):
	
	my_match = re_doxy_bk.match(str)
	if my_match is not None:
		return "///<" + my_match.group(1)
	else:
		return ""

# dumps the given file
def dump( inFR, outFile ):
	for s in inFR:
		outFile.write("."), 
		outFile.write(s)

def processGlobalComments( inFR, outFile ):
	# we have to look for global comments first!
	# they start with '!
	in_block = False
	cnt = 0
	for s in inFR:
		cnt+=1
		if not in_block:
			gcom = re_doxy_class.match(s)
			if gcom is not None:
				# found global comment
				if gcom.group(1) is not None:
					# write this comment to file
					outFile.write("/// " + gcom.group(1) + "\n")
			
			gcom = re_doxy_block_st.match(s)
			if gcom is not None:
				in_block = True
				# found block comment
				if gcom.group(1) is not None:
					# write this comment to file
					outFile.write("/** " + gcom.group(1) + "\n")
		else:
			gcom = re_doxy_block_proc.match(s)
			if gcom is None:
				outFile.write( "*/" )
				break
				
			s = gcom.group(1) + gcom.group(2) + "\n"
			gcom = re_doxy_block_ed.match(s)
			if gcom is not None:
				outFile.write( "*/" )
				break
			
			outFile.write(s)
	else:
		cnt = 0
	
	return cnt

def processClassName( inFR, outFile ):
	sys.stderr.write("Searching for classname... ")
	classBase = None
	className = "dummy"
	for s in inFR:
		if classBase is None:
			cname = re_VB_Obj.match(s)
			if cname is not None:
				classBase = ""
				if cname.group(1) is not None:
					classBase = cname.group(1)
		
		# now search for a class name
		cname = re_VB_Name.match(s)
		if cname is not None:
			# ok, className is found, so save it...
			sys.stderr.write("found!" )
			className = cname.group(1)
			# ...and leave searching-loop
			break
			
	# ok, so let's start writing the pseudo-class
	sys.stderr.write(" using " + className + "\n") 
	
	if classBase is None:
		outFile.write("\nnamespace " + className + "\n{\n") 
	elif classBase is "":
		outFile.write("\nclass " + className + "\n{\n") 
	else:
		outFile.write("\nclass " + className + " : " + classBase + "\n{\n") 

def processFormControl( inFR, outFile ):
	
	ctrls = []
	ctrlpropertys = []
	propertys = []
	outFile.write( "///@name Form Controls\n" )
	outFile.write( "///@{\n" )
	inObj = False
	for s in inFR:
		vb_ctrl = re_VB_Obj_Pr.match(s)
		if vb_ctrl is not None:
			if "Index" == vb_ctrl.group(1): propertys.append(s)
			elif "Caption" == vb_ctrl.group(1): propertys.append(s)
			elif "MaxLength" == vb_ctrl.group(1): propertys.append(s)
			elif "IMEMode" == vb_ctrl.group(1): propertys.append(s)
			elif "Value" == vb_ctrl.group(1): propertys.append(s)
			elif "TabIndex" == vb_ctrl.group(1): propertys.append(s)
			elif "TabStop" == vb_ctrl.group(1): propertys.append(s)
			elif "Enabled" == vb_ctrl.group(1): propertys.append(s)
			elif "Visible" == vb_ctrl.group(1): propertys.append(s)
			elif "WindowList" == vb_ctrl.group(1): propertys.append(s)
			elif "BorderStyle" == vb_ctrl.group(1): propertys.append(s)
			elif "KeyPreview" == vb_ctrl.group(1): propertys.append(s)
			elif "MaxButton" == vb_ctrl.group(1): propertys.append(s)
			elif "StartUpPosition" == vb_ctrl.group(1): propertys.append(s)
			continue
		
		vb_ctrl = re_VB_Obj_Ed.match(s)
		if vb_ctrl is not None:
			if 0 != len(propertys):
				outFile.write( "/**\n@details\t" )
				for pr in propertys:
					outFile.write( "-" + pr )
				outFile.write( "**/\n" )
			outFile.write( ctrls.pop() )
			propertys = ctrlpropertys.pop()
			inObj = False
			continue
		
		vb_ctrl = re_VB_Obj_St.match(s)
		if vb_ctrl is not None:
			ctrls.append( "public:" + vb_ctrl.group(1) + "\t" + vb_ctrl.group(2) + ";\n" )
			ctrlpropertys.append(propertys)
			propertys = []
			continue
		
		if re_VB_Name.match(s) is not None:
			break
			
	outFile.write( "///@}\n" )

# pass blank lines to keep comment block separation
# added by R.S.
def checkBlankLine( outFile, s ):
	global re_blank_line
	
	if re_blank_line.match(s) is None: return False
		
	outFile.write("\n")
	return True

def checkDoxyComment( outFile, s ):
	
	doxy = re_doxy.match(s)
	if doxy is None: return False
	
	outFile.write("/// " + doxy.group(1) + "\n")
	
	return True


# modified by R.S. for const, dim, array, initial value, and so on.
def foundMember( outFile, s ):
	
	member = re_members.match(strip_comments(s))
	if member is None: return False
	
	if member.group(6) is not None:
		#	typeless const declaretion
		initval_str = ""
		if (member.group(7) is not None):
			if (re_const_string.match(member.group(7))):
				initval_str = " = " + re_doublequote.sub(r"\\\"", re_backslash.sub( r"\\\\", member.group(7) ))
			else:
				initval_str = " = " + member.group(7)
		res_str = getAccessibility(member.group(1)) + " const " + member.group(6) + initval_str + ";"
	else:
		#	normal declaretion
		#	check const condition
		const_str = ""
		if (member.group(2) is not None):
			const_str = "const "
		#	check intial value
		initval_str = ""
		if (member.group(5) is not None):
			if (re_const_string.match(member.group(5))):
				initval_str = " = " + re_doublequote.sub(r"\\\"", re_backslash.sub( r"\\\\", member.group(5) ))
			else:
				initval_str = " = " + member.group(5)
		
		#	check array
		valname_str = member.group(3)
		array_idfr = re_array.match(member.group(3))
		if (array_idfr is not None):
			valname_str = array_idfr.group(1) + "[" + array_idfr.group(2) + "]"
		
		#	produce resulting string
		res_str = getAccessibility(member.group(1)) + " " + const_str + " " + (member.group(4) or "") + " " + valname_str + initval_str + ";"
	
	# and deliver it
	outFile.write( res_str + "\t" + doxy_back_comments(s) + "\n" )
	
	return True

# added by R.S.
# modify arglist
def rearrangeArg(argstr):
	
	# get type
	type_str = "Variant"
	if (argstr.group(5) is not None):
		type_str = argstr.group(5)
	# get arg name
	if (argstr.group(4) is not None):
		argname_str = argstr.group(3) + "[]"
	else:
		argname_str = argstr.group(3)
	# get default value
	dfltval_str = ""
	if ((argstr.group(1) is not None) and (argstr.group(6) is not None)):
		if (re_const_string.match(argstr.group(6))):
			dfltval_str = " = " + re_doublequote.sub(r"\\\"", re_backslash.sub(r"\\\\", argstr.group(6) ))
		else:
			dfltval_str = " = " + argstr.group(6)
	return (argstr.group(1) or "") + " " +(argstr.group(2) or "") +" " + type_str + " " + argname_str + " " + dfltval_str

def foundEvent( outFile, s ):
	
	s_event = re_event.match( strip_comments(s) )
	if s_event is None: return False
	
	res_str = getAccessibility( s_event.group(1) ) + " Event " + s_event.group(2) + re_arg.sub( rearrangeArg, s_event.group(3) ) + ";"
	outFile.write( res_str + doxy_back_comments(s) + "\n" )
	
	return True

# modified by R.S. for variant type, and for scan inside function
def foundFunction( outFile, s ):
	
	s_func = re_function.match( strip_comments(s) )	 # s_func == start_of_a_function
	if s_func is None: return False
		
	type_str = "Variant"
	if (s_func.group(5) is not None):
		type_str = s_func.group(5)
	
	# now make the resulting string
	# modified by R.S. to rearrange arglist
	res_str = getAccessibility( s_func.group(1) ) + " " + type_str + " " + s_func.group(3) + re_arg.sub( rearrangeArg, s_func.group(4) ) + "{"
	# and deliver this string
	outFile.write( res_str + "\n" )
	
	return True
	
# added by R.S.	for scan inside function (now, only skip inside)
def processFunction( outFile, s ):
	
	vbEndFunction = re_endFunction.match( strip_comments(s) )
	if vbEndFunction is None: return True
	
	outFile.write("}\n") #write end of function
	
	return False

#  modified by R.S. for check inside sub
def foundSub( outFile, s ):
	
	s_sub = re_sub.match(strip_comments(s))
	if (s_sub is None): return False
	
	#	produce resulting string
	# modified by R.S. to rearrange arglist
	res_str = getAccessibility(s_sub.group(1)) + " Sub " + s_sub.group(3) + re_arg.sub(rearrangeArg, s_sub.group(4))  + "{"
	# and deliver it
	outFile.write(res_str + "\n")
	
	return True

# added by R.S.	for scan inside sub (now, only skip inside)
def processSub( outFile, s ):
	
	vbEndSub = re_endSub.match( strip_comments(s) )
	if (vbEndSub is not None): # found End Sub
		outFile.write("}\n") #write end of function
		return False
		
	else:
		# inside Sub
		return True

def foundProperty( outFile, s ):
	
	s_pro = re_property.match(strip_comments(s))
	if s_pro is None: return False
	
	type_str = ""
	if "Property Get" == s_pro.group(2):
		if s_pro.group(5) is None:
			type_str = "Variant"
		else:
			type_str = s_pro.group(5)
		
	res_str = getAccessibility(s_pro.group(1)) + " " + s_pro.group(2) + " " + type_str +" " + s_pro.group(3) + re_arg.sub(rearrangeArg, s_pro.group(4))  + "{"
	outFile.write(res_str + "\n")
	
	return True

def processProperty( outFile, s ):
	
	vbEndProperty = re_endProperty.match( strip_comments(s) )
	if (vbEndProperty is not None):
		outFile.write("}\n")
		return False
		
	else:
		return True

def foundBlockComment( outFile, s ):
	
	res = re_doxy_block_st.match(s)
	if res is None: return False

	# found block comment
	if res.group(1) is not None:
		# write this comment to file
		outFile.write("/** " + res.group(1) + "\n")
		
	return True

def processBlockComment( outFile, s ):
	
	res = re_doxy_block_proc.match(s)
	if res is None: return False
		
	outFile.write( res.group(1) + res.group(2) + "\n" )
	
	res = re_doxy_block_ed.match(s)
	if res is not None: return False
	
	return True

def getAccessibility(s):
	accessibility = def_level
	if (s is not None):
		if (s.strip().lower() == "private"): accessibility = "private:"
		elif (s.strip().lower() == "public"): accessibility = "public:"
		elif (s.strip().lower() == "friend"): accessibility = "friend "
		elif (s.strip().lower() == "static"): accessibility = "static"
	return accessibility

# modified by R.S. for const, dim, array, initial value, and so on.
def foundMemberOfType( outFile, s ):
	
	member = re_members.match( strip_comments(s) )
	if member is None: return
	
	if member.group(6) is not None:
		#	typeless const declaretion
		initval_str = ""
		if (member.group(7) is not None):
			if (re_const_string.match(member.group(7))):
				initval_str = " = " + re_doublequote.sub( r"\\\"", re_backslash.sub( r"\\\\", member.group(7) ))
			
			else:
				initval_str = " = " + member.group(7)
			
		res_str = "const " + member.group(6) + initval_str + ";"
	
	else:
		#	normal declaretion
		#	check const condition
		const_str = ""
		if (member.group(2) is not None):
			const_str = "const "
		#	check intial value
		initval_str = ""
		if (member.group(5) is not None):
			if (re_const_string.match(member.group(5))):
				initval_str = " = " + re_doublequote.sub( r"\\\"", re_backslash.sub( r"\\\\", member.group(5) ))
			else:
				initval_str = " = " + member.group(5)
		#	check array
		valname_str = member.group(3)
		array_idfr = re_array.match(member.group(3))
		if (array_idfr is not None):
			valname_str = array_idfr.group(1) + "[" + array_idfr.group(2) + "]"
		#	produce resulting string
		res_str = const_str + " " + (member.group(4) or "") + " " + valname_str + initval_str + ";"
	
	# and deliver it
	outFile.write(res_str + doxy_back_comments(s) + "\n")

def foundType( outFile, s ):
	
	vbType = re_type.match( strip_comments(s) )
	if vbType is None: return False
	
	res_str = getAccessibility( vbType.group(1) ) + " struct " + vbType.group(2)  + " {"
	outFile.write( res_str + "\n" )
	return True

def processType( outFile, s ):
	
	vbEndType = re_endType.match( strip_comments(s) )
	if (vbEndType is not None): # found End Type
		outFile.write("}; \n") #write end of struct
		return False
		
	else:
		# match <var AS type>
		# write <type var;>
		foundMemberOfType( outFile, s )
		return True

# modified by R.S. for process enum
def foundEnum( outFile, s ):
	
	vbEnum = re_enum.match(strip_comments(s))
	if vbEnum is None: return False
	
	#	produce resulting string
	res_str = getAccessibility( vbEnum.group(1) ) + " enum " + vbEnum.group(2)  + " {"
	# and deliver it
	outFile.write(res_str + "\n")
	
	return True

# modified by R.S. for process enum
def processEnum( outFile, s ):
	
	vbEndEnum = re_endEnum.match( strip_comments(s) )
	if (vbEndEnum is not None):		# found End Enum
		outFile.write( "}; \n" )	#write end of enum
		return False
	
	else:
		outFile.write(strip_comments(s) + ", " + doxy_back_comments(s) + "\n")
		return True

def filterProgramCode( inFR, outFile, st_line = 0 ):
	
	inSearchFunction = None
	s = None
	hterm = False
	cnt = 0
	
	for ln in inFR:
		if 0 < st_line:
			cnt+=1
			if cnt <= st_line: continue
			st_line = 0
		
		if s is not None:
			ln = s + ln
			s = None
		
		if ((re_comments.match(ln) is None) and (ln[-3:] == " _\n")):
			s = ln[:-2]
			continue
			
		# added by R.S. for pass blank lines to separate each comment block
		checkBlankLine( outFile, ln )
		if checkDoxyComment( outFile, ln ):
			continue
		
		if inSearchFunction is not None:
			if not inSearchFunction( outFile, ln ): inSearchFunction = None
			continue
		
		if foundBlockComment( outFile, ln ):
			inSearchFunction = processBlockComment
			continue
			
		if foundType( outFile, ln ):
			inSearchFunction = processType
			continue
		
		#	see if line contains a member
		#	added by R.S. to proccess variables in BAS file
		if foundMember( outFile, ln ):
			continue # line could not contain anything more than a member
		
		if foundEvent( outFile, ln ):
			continue
		
		# line is not a comment. 
		# see if there is a function-statement
		if foundFunction( outFile, ln ):
			inSearchFunction = processFunction
			continue
		
		# there was no match to a function - let's try a sub
		if foundSub( outFile, ln ):
			inSearchFunction = processSub
			continue
		
		# there was no match to a function - let's try a sub
		if foundProperty( outFile, ln ):
			inSearchFunction = processProperty
			continue
		
		# see if there is an enum declaretion
		if foundEnum( outFile, ln ):
			inSearchFunction = processEnum
			continue

# filters .cls-files - VB-CLASS-FILES
def filterCLS( inFR, outFile ):
	
	outFile.write("\n// -- processed by [filterCLS] --\n") 

	st_line = processGlobalComments( inFR, outFile )
	processClassName( inFR, outFile )
	
	filterProgramCode( inFR, outFile, st_line )
	
	outFile.write("}")
	outFile.write("\n// -- [/filterCLS] --\n") 


# filters .bas-files
def filterBAS( inFR, outFile ):
	
	outFile.write("\n// -- processed by [filterBAS] --\n") 

	st_line = processGlobalComments( inFR, outFile )
	processClassName( inFR, outFile )
	
	filterProgramCode( inFR, outFile, st_line )
	
	outFile.write("}")
	outFile.write("\n// -- [/filterBAS] --\n") 

# filters .frm-files
def filterFRM( inFR, outFile ):
	
	outFile.write("\n// -- processed by [filterFRM] --\n") 
	
	st_line = processGlobalComments( inFR, outFile )
	processClassName( inFR, outFile )
	if optC: processFormControl( inFR, outFile )
	
	filterProgramCode( inFR, outFile, st_line )
	
	outFile.write("}")
	outFile.write("\n// -- [/filterFRM] --\n") 

## main filter-function ##
##
## this function decides whether the file is
## (*) a bas file  - module
## (*) a cls file  - class
## (*) a frm file  - frame
##
## and calls the appropriate function
def filter(inFileName, outFileName=None):
	
	try:
		#output file open
		if outFileName is not None:
			inFile = open( outFileName, encoding = src_encoding )
		else:
			outFile = sys.stdout;
		
		#input file open
		inFile = open( inFileName, encoding = src_encoding )
		inFR = inFile.readlines()
		inFile.close()
		
		root, ext = os.path.splitext(filename)
		
		if		(ext.lower() ==".bas"):	filterBAS( inFR, outFile )	## if it is a module call filterBAS
		elif	(ext.lower() ==".cls"):	filterCLS( inFR, outFile )	## if it is a class call filterCLS
		elif	(ext.lower() == ".frm"):filterFRM( inFR, outFile )	## if it is a frame call filterFRM
		else:	dump( inFR, outFile )								## if it is an unknown extension, just dump it
		
		sys.stderr.write("OK\n")
		
		if outFile is not sys.stdout:
			outFile.close()
		
	except IOError as e:
		sys.stderr.write(e[1]+"\n")

## main-entry ##
################
optC = False
args = len(sys.argv)
if args == 1 or 3 < args:
	print( "usage: ", sys.argv[0], " [option] filename" )
	print( "option: C	Puts Control of Form" )
	sys.exit(1)

# Filter the specified file and print the result to stdout
filename = sys.argv[args-1]
if 2<= args: optC = ("C" == sys.argv[1])
filter(filename)
sys.exit(0)
