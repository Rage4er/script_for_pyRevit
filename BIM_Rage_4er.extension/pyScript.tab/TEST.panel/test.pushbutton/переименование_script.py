# -*- coding: utf-8 -*-

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Drawing import *
import sys
import datetime

# Global variables
enable_logging = True
log_messages = []

# Logging function
def add_log(message):
    global enable_logging
    if not enable_logging:
        return
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    log_message = "[{0}] {1}".format(timestamp, message)
    log_messages.append(log_message)
    print(log_message)

# Show logs
def show_logs():
    global log_messages
    if not enable_logging or not log_messages:
        MessageBox.Show("Logging disabled or no records", "Information")
        return
    log_text = "\n".join(log_messages)
    form = Form()
    form.Text = "Execution Logs"
    form.Size = Size(600, 400)
    textbox = TextBox()
    textbox.Multiline = True
    textbox.Dock = DockStyle.Fill
    textbox.ScrollBars = ScrollBars.Vertical
    textbox.Text = log_text
    form.Controls.Add(textbox)
    form.ShowDialog()

# Rename views function
def rename_views(doc, selected_views, prefix, suffix, start_number, number_format, replace_current_name):
    transaction = Transaction(doc, "Rename Views")
    transaction.Start()
    
    try:
        for i, view in enumerate(selected_views):
            current_name = view.Name
            
            # Number formatting
            if number_format == "1":
                number_str = str(start_number + i)
            elif number_format == "01":
                number_str = str(start_number + i).zfill(2)
            elif number_format == "001":
                number_str = str(start_number + i).zfill(3)
            else:
                number_str = str(start_number + i)
            
            # Form new name
            if replace_current_name:
                new_name = "{0}{1}{2}".format(prefix, number_str, suffix)
            else:
                new_name = "{0}{1}{2}{3}".format(prefix, current_name, number_str, suffix)
            
            # Rename view
            add_log("Attempt to rename view: '{0}' -> '{1}'".format(current_name, new_name))
            view.Name = new_name
            add_log("Successfully renamed view: '{0}' -> '{1}'".format(current_name, new_name))
        
        transaction.Commit()
        add_log("Transaction completed successfully")
        return True, "Successfully renamed {0} views".format(len(selected_views))
    
    except Exception as e:
        transaction.RollBack()
        add_log("Transaction error: {0}".format(str(e)))
        return False, "Error during renaming: {0}".format(str(e))

# Get selected views function
def get_selected_views(uidoc):
    selected_views = []
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()
    
    add_log("Elements received from selection: {0}".format(len(selected_ids)))
    
    for elem_id in selected_ids:
        elem = uidoc.Document.GetElement(elem_id)
        
        if elem is None:
            add_log("Element with ID {0} not found".format(elem_id))
            continue
        
        # SIMPLE CHECK - if element has Name and ViewType properties, it's a view
        if hasattr(elem, 'Name') and hasattr(elem, 'ViewType'):
            selected_views.append(elem)
            add_log("Added view: {0} (ID: {1})".format(elem.Name, elem.Id))
        else:
            add_log("Skipped non-view: {0}".format(elem.GetType().Name))
    
    return selected_views

# Main form
class MainForm(Form):
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.selected_views = []
        self.InitializeComponent()
        add_log("Form initialized")
        
        # Get preselected views automatically
        self.selected_views = get_selected_views(self.uidoc)
        self.UpdateViewsList()
    
    def InitializeComponent(self):
        self.Text = "Batch View Renaming"
        self.Size = Size(750, 550)
        self.StartPosition = FormStartPosition.CenterScreen
        self.tabControl = TabControl()
        self.tabControl.Dock = DockStyle.Fill
        
        tabs = ["1. Select Views", "2. Name Settings", "3. Preview", "4. Execute", "5. Finish"]
        for index, text in enumerate(tabs):
            tab = TabPage()
            tab.Text = text
            if index == 0:
                self.SetupTab1(tab)
            elif index == 1:
                self.SetupTab2(tab)
            elif index == 2:
                self.SetupTab3(tab)
            elif index == 3:
                self.SetupTab4(tab)
            elif index == 4:
                self.SetupTab5(tab)
            self.tabControl.TabPages.Add(tab)
        
        self.Controls.Add(self.tabControl)
        add_log("Form components created")
    
    def CreateControl(self, control_type, **kwargs):
        control = control_type()
        for prop, value in kwargs.items():
            setattr(control, prop, value)
        return control
    
    def SetupTab1(self, tab):
        label = self.CreateControl(Label, Text="Select views for renaming", Location=Point(10, 10), Size=Size(300, 20), Font=Font("Arial", 10, FontStyle.Bold))
        self.btnRefreshSelection = self.CreateControl(Button, Text="Refresh Selection", Location=Point(10, 40), Size=Size(120, 25))
        self.btnAddToSelection = self.CreateControl(Button, Text="Add to Selection", Location=Point(140, 40), Size=Size(120, 25))
        self.listViews = self.CreateControl(ListBox, Location=Point(10, 80), Size=Size(700, 300))
        self.lblSelectedCount = self.CreateControl(Label, Text="Selected views: 0", Location=Point(10, 390), Size=Size(150, 20))
        self.btnNext1 = self.CreateControl(Button, Text="Next", Location=Point(600, 450), Size=Size(80, 25))
        
        self.btnRefreshSelection.Click += self.OnRefreshSelectionClick
        self.btnAddToSelection.Click += self.OnAddToSelectionClick
        self.btnNext1.Click += self.OnNext1Click
        
        tab.Controls.Add(label)
        tab.Controls.Add(self.btnRefreshSelection)
        tab.Controls.Add(self.btnAddToSelection)
        tab.Controls.Add(self.listViews)
        tab.Controls.Add(self.lblSelectedCount)
        tab.Controls.Add(self.btnNext1)
        
        add_log("Tab 1 configured")
    
    def SetupTab2(self, tab):
        label = self.CreateControl(Label, Text="Name settings", Location=Point(10, 10), Size=Size(300, 20), Font=Font("Arial", 10, FontStyle.Bold))
        
        label_prefix = self.CreateControl(Label, Text="Prefix:", Location=Point(10, 50), Size=Size(100, 20))
        self.txtPrefix = self.CreateControl(TextBox, Location=Point(120, 50), Size=Size(150, 20), Text="")
        
        label_suffix = self.CreateControl(Label, Text="Suffix:", Location=Point(10, 80), Size=Size(100, 20))
        self.txtSuffix = self.CreateControl(TextBox, Location=Point(120, 80), Size=Size(150, 20), Text="")
        
        label_start = self.CreateControl(Label, Text="Start number:", Location=Point(10, 110), Size=Size(100, 20))
        self.numStartNumber = self.CreateControl(NumericUpDown, Location=Point(120, 110), Size=Size(80, 20), Minimum=1, Maximum=1000, Value=1)
        
        label_format = self.CreateControl(Label, Text="Number format:", Location=Point(10, 140), Size=Size(100, 20))
        self.cmbNumberFormat = self.CreateControl(ComboBox, Location=Point(120, 140), Size=Size(80, 20))
        
        self.chkReplaceName = self.CreateControl(CheckBox, Text="Replace current name", Location=Point(10, 170), Size=Size(200, 20), Checked=False)
        
        self.btnBack1 = self.CreateControl(Button, Text="Back", Location=Point(500, 450), Size=Size(80, 25))
        self.btnNext2 = self.CreateControl(Button, Text="Next", Location=Point(600, 450), Size=Size(80, 25))
        
        self.cmbNumberFormat.Items.Add("1")
        self.cmbNumberFormat.Items.Add("01")
        self.cmbNumberFormat.Items.Add("001")
        self.cmbNumberFormat.SelectedIndex = 0
        
        self.btnBack1.Click += self.OnBack1Click
        self.btnNext2.Click += self.OnNext2Click
        
        tab.Controls.Add(label)
        tab.Controls.Add(label_prefix)
        tab.Controls.Add(self.txtPrefix)
        tab.Controls.Add(label_suffix)
        tab.Controls.Add(self.txtSuffix)
        tab.Controls.Add(label_start)
        tab.Controls.Add(self.numStartNumber)
        tab.Controls.Add(label_format)
        tab.Controls.Add(self.cmbNumberFormat)
        tab.Controls.Add(self.chkReplaceName)
        tab.Controls.Add(self.btnBack1)
        tab.Controls.Add(self.btnNext2)
        
        add_log("Tab 2 configured")
    
    def SetupTab3(self, tab):
        label = self.CreateControl(Label, Text="Preview changes", Location=Point(10, 10), Size=Size(300, 20), Font=Font("Arial", 10, FontStyle.Bold))
        self.listPreview = self.CreateControl(ListBox, Location=Point(10, 40), Size=Size(700, 350))
        self.btnUpdatePreview = self.CreateControl(Button, Text="Update Preview", Location=Point(10, 400), Size=Size(150, 25))
        self.btnBack2 = self.CreateControl(Button, Text="Back", Location=Point(500, 450), Size=Size(80, 25))
        self.btnNext3 = self.CreateControl(Button, Text="Next", Location=Point(600, 450), Size=Size(80, 25))
        
        self.btnUpdatePreview.Click += self.OnUpdatePreviewClick
        self.btnBack2.Click += self.OnBack2Click
        self.btnNext3.Click += self.OnNext3Click
        
        tab.Controls.Add(label)
        tab.Controls.Add(self.listPreview)
        tab.Controls.Add(self.btnUpdatePreview)
        tab.Controls.Add(self.btnBack2)
        tab.Controls.Add(self.btnNext3)
        
        add_log("Tab 3 configured")
    
    def SetupTab4(self, tab):
        label = self.CreateControl(Label, Text="Execute renaming", Location=Point(10, 10), Size=Size(300, 20), Font=Font("Arial", 10, FontStyle.Bold))
        self.btnExecute = self.CreateControl(Button, Text="Execute Renaming", Location=Point(10, 40), Size=Size(180, 30), BackColor=Color.LightGreen)
        self.txtResults = self.CreateControl(TextBox, Location=Point(10, 80), Size=Size(700, 320), Multiline=True, ReadOnly=True, ScrollBars=ScrollBars.Vertical)
        self.btnShowLogs2 = self.CreateControl(Button, Text="Show Logs", Location=Point(400, 450), Size=Size(100, 25))
        self.btnBack3 = self.CreateControl(Button, Text="Back", Location=Point(500, 450), Size=Size(80, 25))
        self.btnNext4 = self.CreateControl(Button, Text="Next", Location=Point(600, 450), Size=Size(80, 25))
        
        self.btnExecute.Click += self.OnExecuteClick
        self.btnShowLogs2.Click += self.OnShowLogsClick
        self.btnBack3.Click += self.OnBack3Click
        self.btnNext4.Click += self.OnNext4Click
        
        tab.Controls.Add(label)
        tab.Controls.Add(self.btnExecute)
        tab.Controls.Add(self.txtResults)
        tab.Controls.Add(self.btnShowLogs2)
        tab.Controls.Add(self.btnBack3)
        tab.Controls.Add(self.btnNext4)
        
        add_log("Tab 4 configured")
    
    def SetupTab5(self, tab):
        label = self.CreateControl(Label, Text="Renaming completed", Location=Point(10, 10), Size=Size(300, 20), Font=Font("Arial", 10, FontStyle.Bold))
        label_info = self.CreateControl(Label, Text="View renaming operation completed successfully.", Location=Point(10, 40), Size=Size(500, 20))
        self.btnShowLogs = self.CreateControl(Button, Text="Show Logs", Location=Point(10, 70), Size=Size(100, 25))
        self.btnBack4 = self.CreateControl(Button, Text="Back", Location=Point(500, 450), Size=Size(80, 25))
        self.btnFinish = self.CreateControl(Button, Text="Finish", Location=Point(600, 450), Size=Size(80, 25))
        
        self.btnShowLogs.Click += self.OnShowLogsClick
        self.btnBack4.Click += self.OnBack4Click
        self.btnFinish.Click += self.OnFinishClick
        
        tab.Controls.Add(label)
        tab.Controls.Add(label_info)
        tab.Controls.Add(self.btnShowLogs)
        tab.Controls.Add(self.btnBack4)
        tab.Controls.Add(self.btnFinish)
        
        add_log("Tab 5 configured")
    
    def UpdateViewsList(self):
        self.listViews.Items.Clear()
        for view in self.selected_views:
            self.listViews.Items.Add("{0} (ID: {1})".format(view.Name, view.Id))
        
        self.lblSelectedCount.Text = "Selected views: {0}".format(len(self.selected_views))
        add_log("View list updated. Selected views: {0}".format(len(self.selected_views)))
    
    def OnRefreshSelectionClick(self, sender, args):
        try:
            add_log("Refreshing selection...")
            self.selected_views = get_selected_views(self.uidoc)
            self.UpdateViewsList()
            
            if len(self.selected_views) == 0:
                MessageBox.Show("No views selected! Select views in Revit and click the button again.", "Warning")
                add_log("Warning: no views selected")
        
        except Exception as e:
            error_msg = "Error refreshing selection: {0}".format(str(e))
            add_log(error_msg)
            MessageBox.Show(error_msg, "Error")
    
    def OnAddToSelectionClick(self, sender, args):
        try:
            add_log("Adding to selection...")
            new_views = get_selected_views(self.uidoc)
            
            existing_ids = [v.Id for v in self.selected_views]
            for view in new_views:
                if view.Id not in existing_ids:
                    self.selected_views.append(view)
                    add_log("Added view: {0} (ID: {1})".format(view.Name, view.Id))
            
            self.UpdateViewsList()
            add_log("Addition completed. Total views: {0}".format(len(self.selected_views)))
            
        except Exception as e:
            error_msg = "Error adding to selection: {0}".format(str(e))
            add_log(error_msg)
            MessageBox.Show(error_msg, "Error")
    
    def OnUpdatePreviewClick(self, sender, args):
        add_log("Updating preview...")
        if not self.selected_views:
            MessageBox.Show("First select views on tab 1", "Warning")
            add_log("Warning: no views selected for preview")
            return
        
        prefix = self.txtPrefix.Text
        suffix = self.txtSuffix.Text
        start_number = int(self.numStartNumber.Value)
        number_format = self.cmbNumberFormat.SelectedItem.ToString()
        replace_current = self.chkReplaceName.Checked
        
        add_log("Renaming parameters: prefix='{0}', suffix='{1}', start={2}, format={3}, replace={4}".format(prefix, suffix, start_number, number_format, replace_current))
        
        self.listPreview.Items.Clear()
        
        for i, view in enumerate(self.selected_views):
            current_name = view.Name
            
            if number_format == "1":
                number_str = str(start_number + i)
            elif number_format == "01":
                number_str = str(start_number + i).zfill(2)
            elif number_format == "001":
                number_str = str(start_number + i).zfill(3)
            else:
                number_str = str(start_number + i)
            
            if replace_current:
                new_name = "{0}{1}{2}".format(prefix, number_str, suffix)
            else:
                new_name = "{0}{1}{2}{3}".format(prefix, current_name, number_str, suffix)
            
            self.listPreview.Items.Add("{0} -> {1}".format(current_name, new_name))
        
        add_log("Preview updated for {0} views".format(len(self.selected_views)))
    
    def OnExecuteClick(self, sender, args):
        add_log("Starting renaming execution...")
        if not self.selected_views:
            MessageBox.Show("No views selected", "Error")
            add_log("Error: no views selected for renaming")
            return
        
        prefix = self.txtPrefix.Text
        suffix = self.txtSuffix.Text
        start_number = int(self.numStartNumber.Value)
        number_format = self.cmbNumberFormat.SelectedItem.ToString()
        replace_current = self.chkReplaceName.Checked
        
        self.txtResults.Text = "Starting renaming...\n"
        add_log("Starting renaming of {0} views".format(len(self.selected_views)))
        
        success, message = rename_views(self.doc, self.selected_views, prefix, suffix, start_number, number_format, replace_current)
        
        self.txtResults.Text += message + "\n"
        add_log("Renaming result: {0}".format(message))
        
        if success:
            self.txtResults.Text += "Renaming completed successfully!"
            add_log("Renaming completed successfully")
        else:
            self.txtResults.Text += "Error occurred during renaming."
            add_log("Renaming completed with errors")
    
    def OnShowLogsClick(self, sender, args):
        add_log("Log display requested by user")
        show_logs()
    
    def OnNext1Click(self, sender, args):
        if len(self.selected_views) == 0:
            MessageBox.Show("First select views", "Warning")
            add_log("Warning: moving to tab 2 without views")
            return
        self.tabControl.SelectedIndex = 1
        add_log("Moving to tab 2")
    
    def OnNext2Click(self, sender, args):
        self.tabControl.SelectedIndex = 2
        add_log("Moving to tab 3")
        self.OnUpdatePreviewClick(None, None)
    
    def OnNext3Click(self, sender, args):
        self.tabControl.SelectedIndex = 3
        add_log("Moving to tab 4")
    
    def OnNext4Click(self, sender, args):
        self.tabControl.SelectedIndex = 4
        add_log("Moving to tab 5")
    
    def OnBack1Click(self, sender, args):
        self.tabControl.SelectedIndex = 0
        add_log("Moving to tab 1")
    
    def OnBack2Click(self, sender, args):
        self.tabControl.SelectedIndex = 1
        add_log("Moving to tab 2")
    
    def OnBack3Click(self, sender, args):
        self.tabControl.SelectedIndex = 2
        add_log("Moving to tab 3")
    
    def OnBack4Click(self, sender, args):
        self.tabControl.SelectedIndex = 3
        add_log("Moving to tab 4")
    
    def OnFinishClick(self, sender, args):
        add_log("Closing application")
        MessageBox.Show("Operation completed!")
        self.Close()

def main():
    try:
        add_log("Starting script...")
        doc = __revit__.ActiveUIDocument.Document
        uidoc = __revit__.ActiveUIDocument
        add_log("Revit document acquired")
        
        if doc and uidoc:
            add_log("Starting main form")
            Application.Run(MainForm(doc, uidoc))
        else:
            error_msg = "No access to Revit document"
            add_log(error_msg)
            MessageBox.Show(error_msg)
    except Exception as e:
        error_msg = "Error: " + str(e)
        add_log(error_msg)
        MessageBox.Show(error_msg)

if __name__ == "__main__":
    main()