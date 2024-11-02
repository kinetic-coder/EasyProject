import win32com.client
import os

class MicrosoftProjectPlan:
    '''
    Represents a Microsoft Project Plan 
    '''
    def __init__(self, name, start_date, end_date, path):
        self.name = name
        self.start_date = start_date
        self.end_date = end_date
        self.path = path
        
    def get_duration(self):
        return (self.end_date - self.start_date).days

    def __str__(self):
        return f"Project: {self.name}, Start Date: {self.start_date}, End Date: {self.end_date}, Duration: {self.get_duration()} days"
    
    '''
    Opens or creates a new Micoroft Project Plan file located on disk.
    '''
    def open(self):
        try:
            if os.path.exists(self.path):
                ms_project = win32com.client.Dispatch("MSProject.Application")
                ms_project.FileOpen(path)
            else:
                ms_project = win32com.client.Dispatch("MSProject.Application")
                ms_project.FileNew()
                ms_project.FileSaveAs(self.path)
            
            ms_project.Visible = True
            project = ms_project.FileOpen(self.path)
            return ms_project
        except Exception as e:
            print(f"Failed to open project: {e}")
            return None