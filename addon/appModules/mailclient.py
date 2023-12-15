# -*- coding: UTF-8 -*-
#A part of em Client addon for NVDA
#Copyright (C) 2020 Tony Malykh
#This file is covered by the GNU General Public License.
#See the file COPYING.txt for more details.

import addonHandler
import api
import appModuleHandler
import config
from collections import defaultdict
import controlTypes
from NVDAObjects.behaviors import NVDAObject, RowWithFakeNavigation
from NVDAObjects.UIA import UIA
from scriptHandler import script
import speech
import tones
import ui
import UIAHandler
from UIAHandler.utils import createUIAMultiPropertyCondition

addonHandler.initTranslation()

# Translators: corresponding column in Em Client
COL_READ_STATUS = _("Read status")
# Translators: corresponding column in Em Client
COL_FLAG = _("Flag")
# Translators: corresponding column in Em Client
COL_FROM = _("From")
# Translators: corresponding column in Em Client
COL_SUBJECT = _("Subject")
# Translators: corresponding column in Em Client
COL_RECEIVED = _("Received")
# Translators: corresponding column in Em Client
COL_SIZE = _("Size")
# Translators: corresponding column in Em Client
COL_MAIL_FOLDER = _("Mail Folder")
# Translators: corresponding column in Em Client
COL_ATTACHMENT = _("Attachment")
# Translators: corresponding column in Em Client
COL_ACCOUNT_ICON = _("Account Icon")
# Translators: corresponding column in Em Client
COL_DELETE = _("Delete")


def printTree3(obj, level=10, indent=0):
    result = []
    indentStr = " "*indent
    if obj is None:
        return f"{indentStr}<None>"
    if level < 0:
        return f"{indentStr}..."
    name = obj.name or "<None>"
    if "\n" in name:
        indentStr2 = indentStr + "    "
        name = "\n" + "\n".join(
            [
                indentStr2 + s
                for s in name.split("\n")
            ]
        )
    desc = f"{indentStr}{controlTypes.roleLabels[obj.role]} {name}"
    result.append(desc)
    ni = indent+4
    li = level-1
    child = obj.simpleFirstChild
    while child is not None:
        result.append(printTree3(child, li, ni))
        child = child.simpleNext
    return "\n".join(result)

def getWindow(focus):
    if focus.parent is None:
        raise Exception("Desktop window is focused!")
    while focus.parent.parent is not None:
        focus = focus.parent
    return focus


def findDocument(window=None):
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

class EmClientBase:

    def getChildren(self, obj=None):
        if obj is None:
            obj = self
        # Collecting all children as a single request in order to make this real fast - code adopted from Outlook appModule
        childrenCacheRequest=UIAHandler.handler.baseCacheRequest.clone()
        childrenCacheRequest.addProperty(UIAHandler.UIA_NamePropertyId)
        childrenCacheRequest.addProperty(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId)
        childrenCacheRequest.TreeScope=UIAHandler.TreeScope_Children
        # We must filter the children for just text and image elements otherwise getCachedChildren fails completely in conversation view.
        childrenCacheRequest.treeFilter=createUIAMultiPropertyCondition({UIAHandler.UIA_ControlTypePropertyId:[UIAHandler.UIA_TextControlTypeId,UIAHandler.UIA_ImageControlTypeId]})
        cachedChildren=obj.UIAElement.buildUpdatedCache(childrenCacheRequest).getCachedChildren()
        return  cachedChildren


class MailViewRow(RowWithFakeNavigation,UIA,EmClientBase):

    def composeName(self):

        result = []

        if controlTypes.State.EXPANDED in self.states:
            result.append(controlTypes.State.EXPANDED.displayString)
        elif controlTypes.State.COLLAPSED in self.states:
            result.append(controlTypes.State.COLLAPSED.displayString)

        cachedChildren = self.getChildren()

        for i in range(cachedChildren.length):
            child = cachedChildren.getElement(i)

            if child.cachedName == '':
                continue

            headerText = child.getCachedPropertyValueEx(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,True).QueryInterface(UIAHandler.IUIAutomationElementArray).getElement(0).CurrentName

            if headerText == COL_READ_STATUS:
                # translators: text indicating an unread mail, official translation from eM Client must be used here
                t = _("Unread")
                if child.cachedName == t:
                    result.append(t)
                else:
                    continue
            else:
                if child.cachedName == "True":
                    result.append(headerText)
                elif child.cachedName == 'False':
                    continue
                else:
                    result.append(headerText + ": " + child.cachedName)

        return ", ".join(result)

    def _get_name(self):

        if not hasattr(self, '_custom_name'):
            self._custom_name = self.composeName()

        return self._custom_name

    value = None

    def reportFocus(self):
        speech.speakText(self.name)

    @script(gestures=["kb:rightArrow"])
    def script_rightArrow(self, gesture):
        tones.beep(200, 50)

    def _get_role(self):
        role=super(MailViewRow, self).role
        if role==controlTypes.Role.DATAITEM:
            role=controlTypes.Role.LISTITEM
        return role

    def findNextUnread(self, direction, errorMsg):
        readStatusIndex = -1
        cachedChildren = self.getChildren()
        for index in range(cachedChildren.length):
            child = cachedChildren.GetElement(index)
            columnHeaderText = child.getCachedPropertyValueEx(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,True).QueryInterface(UIAHandler.IUIAutomationElementArray).getElement(0).CurrentName
            if columnHeaderText == self.readStatus:
                readStatusIndex = index
                break
        if readStatusIndex < 0:
            raise Exception("Could not find read status column")
        obj = self
        for i in range(1000):
            obj = obj.next if direction > 0 else obj.previous
            if obj is None:
                ui.message(errorMsg)
                return
            cachedChildren = self.getChildren(obj)
            name = cachedChildren.GetElement(readStatusIndex).CachedName
            if name == "False":
                obj.setFocus()
                return
        raise Exception("Could not find unread email after 1000 iterations!")

    @script(description='Find next unread email', gestures=['kb:N'])
    def script_nextUnread(self, gesture):
        return self.findNextUnread(1, _("No next unread email"))

    @script(description='Find previous unread email', gestures=['kb:P'])
    def script_previousUnread(self, gesture):
        return self.findNextUnread(-1, _("No next previous email"))

    """
    def _get_previous(self):
        prev = super()._get_previous()
        if prev is not None:
            return prev
        parent = self.parent
        counter = 0
        while parent.previous is not None and parent.role == parent.previous.role:
            parent = parent.previous
            if len(parent.children) > 0:
                if counter > 0:
                    # We must have skipped over a collapsed section. Beep to let user know.
                    tones.beep(200, 50)
                return parent.children[-1]
            counter += 1
        return None

    def _get_next(self):
        next = super()._get_next()
        if next is not None:
            return next
        parent = self.parent
        counter = 0
        while parent.next is not None and parent.role == parent.next.role:
            parent = parent.next
            if len(parent.children) > 0:
                if counter > 0:
                    # We must have skipped over a collapsed section. Beep to let user know.
                    tones.beep(200, 50)
                return parent.children[0]
            counter  += 1
        return None
    """

    @script(description='Read current email message.', gestures=['kb:NVDA+DownArrow'])
    def script_readEmail(self, gesture):
        document = findSubDocument()
        speakObject(document)


class SettingsViewRow(RowWithFakeNavigation,UIA,EmClientBase):

    def _get_name(self):
        result = []

        if controlTypes.State.EXPANDED in self.states:
            result.append(controlTypes.State.EXPANDED.displayString)
        elif controlTypes.State.COLLAPSED in self.states:
            result.append(controlTypes.State.COLLAPSED.displayString)

        result.append(super(SettingsViewRow, self)._get_name())

        return ", ".join(result)

    value = None

    def reportFocus(self):
        speech.speakText(self.name)

    def _get_role(self):
        role=super(SettingsViewRow, self).role
        if role==controlTypes.Role.DATAITEM:
            role=controlTypes.Role.LISTITEM
        return role


class MenuItem(UIA):

    description = ''


class AppModule(appModuleHandler.AppModule):

    def chooseNVDAObjectOverlayClasses(self, obj, clsList):

        def checkParent(obj, level, uiaAutomationId):
            current = 0
            
            while obj.parent and current < level:
                obj = obj.parent
                current += 1
                
            if current < level:
                return False

            try:
                return obj.UIAAutomationId == uiaAutomationId
            except AttributeError:
                return False

        if isinstance(obj, UIA):
            if obj.role == controlTypes.Role.DATAITEM:
                if checkParent(obj, 2, "dataGridCategory"):
                    clsList.insert(0, SettingsViewRow)
                elif checkParent(obj, 2, "controlDataGrid"):
                    clsList.insert(0, MailViewRow)
            elif obj.role == controlTypes.Role.MENUITEM or (obj.role == controlTypes.Role.PANE and obj.parent.role == controlTypes.Role.POPUPMENU):
                clsList.insert(0, MenuItem)

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
        ui.message(_("Expanded%d messages") % len(headings))
