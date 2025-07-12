#***************************************************************************
#*    Copyright (C) 2023 
#*    This library is free software
#***************************************************************************
import inspect
import os
import sys
import FreeCAD
import FreeCADGui

class DustRemovalShowCommand:
    def GetResources(self):
        file_path = inspect.getfile(inspect.currentframe())
        module_path=os.path.dirname(file_path)
        return { 
          'Pixmap': os.path.join(module_path, "icons", "dustRemoval.svg"),
          'MenuText': "dustRemoval",
          'ToolTip': "Show/Hide dustRemoval"}

    def IsActive(self):
        import DustRemovalAssy
        DustRemovalAssy
        return True

    def Activated(self):
        try:
          import DustRemovalAssy
          DustRemovalAssy.main.d.show()
        except Exception as e:
          FreeCAD.Console.PrintError(str(e) + "\n")

    def IsActive(self):
        import DustRemovalAssy
        return not FreeCAD.ActiveDocument is None

class DustRemoval(FreeCADGui.Workbench):
    def __init__(self):
        file_path = inspect.getfile(inspect.currentframe())
        module_path=os.path.dirname(file_path)
        self.__class__.Icon = os.path.join(module_path, "icons", "dustRemoval.svg")
        self.__class__.MenuText = "DustRemoval"
        self.__class__.ToolTip = "DustRemoval by Pascal"

    def Initialize(self):
        self.commandList = ["DustRemovalShow"]
        self.appendToolbar("&DustRemoval", self.commandList)
        self.appendMenu("&DustRemoval", self.commandList)

    def Activated(self):
        import DustRemovalAssy
        DustRemovalAssy
        return

    def Deactivated(self):
        return

    def ContextMenu(self, recipient):
        return

    def GetClassName(self): 
        return "Gui::PythonWorkbench"
FreeCADGui.addWorkbench(DustRemoval())
FreeCADGui.addCommand("DustRemovalShow", DustRemovalShowCommand())

