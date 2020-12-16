# -*- coding: UTF-8 -*-
#A part of em Client addon for NVDA
#Copyright (C) 2020 Tony Malykh
#This file is covered by the GNU General Public License.
#See the file COPYING.txt for more details.

import addonHandler
import api
import appModuleHandler
import bisect
import config
import controlTypes
import ctypes
import eventHandler
import globalPluginHandler
import gui
import json
import NVDAHelper
from NVDAObjects.behaviors import RowWithFakeNavigation, Dialog
from NVDAObjects.UIA import UIA
from NVDAObjects.window import winword
import operator
import re
import sayAllHandler
from scriptHandler import script, willSayAllResume
import speech
import struct
import textInfos
import time
import tones
import ui
import UIAHandler
from UIAUtils import createUIAMultiPropertyCondition
import wx

debug = False
if debug:
    f = open("C:\\Users\\tony\\Dropbox\\1.txt", "w", encoding="utf-8")
def mylog(s):
    if debug:
        print(str(s), file=f)
        f.flush()

def myAssert(condition):
    if not condition:
        raise RuntimeError("Assertion failed")

#useful for debug!
def printTree(obj, level=10, indent=0):
    result = []
    indentStr = " "*indent
    if level < 0:
        return [f"{indentStr}..."]
    try:
        desc = f"{indentStr}{controlTypes.roleLabels[obj.role]} {obj.name}"
    except:
        desc = str(type(obj))
    result.append(desc)
    ni = indent+4
    li = level-1
    try:
        children = obj.children
    except:
        children = []
    for child in children:
        result.extend(printTree(child, li, ni))
    return "\n".join(result)

def printTree2(obj, level=10, indent=0):
    result = []
    indentStr = " "*indent
    if level < 0:
        return f"{indentStr}..."
    desc = f"{indentStr}{controlTypes.roleLabels[obj.role]} {obj.name}"
    result.append(desc)
    ni = indent+4
    li = level-1
    child = obj.firstChild
    while child is not None:
        result.append(printTree2(child, li, ni))
        child = child.next
    return "\n".join(result)

def printTree3(obj, level=10, indent=0):
    result = []
    indentStr = " "*indent
    if obj is None:
        return f"{indentStr}<None>"
    if level < 0:
        return f"{indentStr}..."
    desc = f"{indentStr}{controlTypes.roleLabels[obj.role]} {obj.name}"
    result.append(desc)
    ni = indent+4
    li = level-1
    child = obj.simpleFirstChild
    while child is not None:
        result.append(printTree3(child, li, ni))
        child = child.simpleNext
    return "\n".join(result)

    
def desc(obj):
    return f"{controlTypes.roleLabels[obj.role]} {obj.name}"
    
def getWindow(focus):
    if focus.parent is None:
        raise Exception("Desktop window is focused!")
    while focus.parent.parent is not None:
        focus = focus.parent
    return focus
def     findDocument(window=None):
    if window is None:
        window = api.getForegroundObject()
    document = window.simpleFirstChild.simpleNext
    if document.role != controlTypes.ROLE_DOCUMENT:
        raise Exception(f"Failed to find document. Debug:\n{printTree3(window)}")
    return document

def findSubDocument(window=None):
    document = findDocument(window)
    subdocument = document.simpleLastChild
    if subdocument.role != controlTypes.ROLE_DOCUMENT:
        subdocument = subdocument.simplePrevious
    if subdocument.role != controlTypes.ROLE_DOCUMENT:
        raise Exception(f"Failed to find subdocument. Debug:\n{printTree3(document)}")
    return subdocument
    
def findTopLevelObject(focus=None, window=None):
    if window is None:
        window = api.getForegroundObject()
    if focus is None:
        focus = api.getFocusObject()
    while focus.parent is not None:
        if focus.simpleParent == window:
            return focus
        focus = focus.simpleParent
    raise Exception("Something went wrong!")

def circularSimpleNext(obj, direction):
    next = obj.simpleNext if direction > 0 else obj.simplePrevious
    if next is None:
        next = obj.simpleParent.simpleFirstChild if direction > 0 else obj.simpleParent.simpleLastChild
    return next
    
def findNextImportant(direction, focus=None, window=None):
    tlo = findTopLevelObject(focus, window)
    obj = tlo
    for i in range(100):
        obj = circularSimpleNext(obj, direction)
        if obj == tlo:
            raise Exception(f"Failed to find next top-level object. Debug:\n{printTree3(window)}")
        if obj.role in {
            controlTypes.ROLE_TABLE,
            controlTypes.ROLE_DOCUMENT,
            controlTypes.ROLE_TREEVIEW,
        } :
            return obj
    raise Exception(f"Failed to find next top-level object - infinite loop detected. Debug:\n{printTree3(window)}")
        

def traverseText(obj):
    child = obj.simpleFirstChild
    if child is None and obj.name is not None and len(obj.name) > 0:
        yield obj.name
    while child is not None:
        for s in traverseText(child):
            yield s
        child = child.simpleNext


def speakObject(document):
    # Try also using:
    # sayAllHandler.readObjects(document)
    generator = traverseText(document)
    def callback():
        try:
            text = generator.__next__()
        except StopIteration:
            return
        speech.speak([text, speech.commands.CallbackCommand(callback)])

    callback()


class AppModule(appModuleHandler.AppModule):
    def chooseNVDAObjectOverlayClasses(self, obj, clsList):
        if obj.role == controlTypes.ROLE_LISTITEM:
            if obj.parent is not None and obj.parent.parent is not None and obj.parent.parent.role == controlTypes.ROLE_TABLE:
                clsList.insert(0, UIAGridRow)

    @script(description='Expand all messages in message view', gestures=['kb:NVDA+X'])
    def script_expandMessages(self, gesture):
        focus = api.getFocusObject()
        interceptor = focus.treeInterceptor
        if interceptor is None:
            ui.message(_("Not in message view!"))
            return
        headings = list(interceptor._iterNodesByType("heading2"))
        for heading in headings:
            if heading.obj.IA2Attributes.get('class', "") == "header header_gray":
                heading.obj.doAction()
        ui.message(_("Expanded"))
        ui.message(f"Found {len(headings)} headings")
        
    @script(description='Jump to next pane', gestures=['kb:F6'])
    def script_nextPane(self, gesture):
        obj = findNextImportant(1)
        obj.setFocus()
        api.setFocusObject(obj)
    @script(description='Jump to previous pane', gestures=['kb:Shift+F6'])
    def script_previousPane(self, gesture):
        obj = findNextImportant(-1)
        obj.setFocus()
        api.setFocusObject(obj)
    def event_gainFocus(self, obj, nextHandler):
        nextHandler()
    def event_focusEntered(self,obj,nextHandler):
        nextHandler()
    def event_UIA_window_windowOpen(self, obj, nextHandler):
        eventHandler.executeEvent("gainFocus", obj)
        # We don't use sayAllHandler.readObjects(obj) here, since it would read the title of the window again.
        speakObject(obj)
        nextHandler()
    
    
class UIAGridRow(RowWithFakeNavigation,UIA):
    def _get_name(self):
        return ""
    def _get_value(self):
        result = []
        # Collecting all children as a single request in order to make this real fast - code adopted from Outlook appModule
        childrenCacheRequest=UIAHandler.handler.baseCacheRequest.clone()
        childrenCacheRequest.addProperty(UIAHandler.UIA_NamePropertyId)
        childrenCacheRequest.addProperty(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId)
        childrenCacheRequest.TreeScope=UIAHandler.TreeScope_Children
        # We must filter the children for just text and image elements otherwise getCachedChildren fails completely in conversation view.
        childrenCacheRequest.treeFilter=createUIAMultiPropertyCondition({UIAHandler.UIA_ControlTypePropertyId:[UIAHandler.UIA_TextControlTypeId,UIAHandler.UIA_ImageControlTypeId]})
        cachedChildren=self.UIAElement.buildUpdatedCache(childrenCacheRequest).getCachedChildren()

        for index in range(cachedChildren.length):
            child = cachedChildren.getElement(index)
            name = child.CurrentName
            if child.cachedControlType == UIAHandler.UIA_ImageControlTypeId:
                # I adore the beauty of COM interfaces!
                columnHeaderText = child.getCachedPropertyValueEx(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,True).QueryInterface(UIAHandler.IUIAutomationElementArray).getElement(0).CurrentName
                name = child.CurrentName
                if columnHeaderText == "Read status":
                    if name == "False":
                        result.append("Unread")
                    elif name == "True":
                        pass
                    else:
                        result.append(columnHeaderText + ": " + name)
                elif name != "False":
                    if name == "True":
                        result.append(columnHeaderText)
                    else:
                        result.append(columnHeaderText + ": " + name)
            else:
                result.append(name)
        return " ".join(result)
    def _get_previous(self):
        prev = super()._get_previous()
        if prev is not None:
            return prev
        parent = self.parent.previous
        while parent is not None:
            if len(parent.children) > 0:
                return parent.children[-1]
            parent = parent.previous


    def _get_next(self):
        next = super()._get_next()
        if next is not None:
            return next
        parent = self.parent.next
        while parent is not None:
            if len(parent.children) > 0:
                return parent.children[0]
            parent = parent.next
    @script(description='Read current email message.', gestures=['kb:NVDA+DownArrow'])
    def script_readEmail(self, gesture):
        document = findSubDocument()
        speakObject(document)
