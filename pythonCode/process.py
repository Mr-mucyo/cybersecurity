import wmi
import pandas as pd 

# Connect to WMI
conn = wmi.WMI()

process_data = []

# Iterate over running processes
for process in conn.Win32_Process():
    
    process_info = {
        "Process ID": process.ProcessId,
        "Handle Count": process.HandleCount,
        "Name": process.Name,
        "Creation Date": process.CreationDate,
        "Thread Count": process.ThreadCount,
        "Virtual Size": process.VirtualSize,
        "Working Set Size": process.WorkingSetSize,
        "Priority": process.Priority,
        "Parent Process ID": process.ParentProcessId,
        "Peak Virtual Size": process.PeakVirtualSize,
        "Peak Working Set Size": process.PeakWorkingSetSize,
        "Page Faults": process.PageFaults,
        "Page File Usage": process.PageFileUsage,
        "Peak Page File Usage": process.PeakPageFileUsage,
        "User Mode Time": process.UserModeTime,
        "Kernel Mode Time": process.KernelModeTime,
        "Windows Version": process.WindowsVersion,
        "Write Operation Count": process.WriteOperationCount,
        "Install Date": process.InstallDate,
        "Read Transfer Count": process.ReadTransferCount,
        "Termination Date": process.TerminationDate,
        "Quota Non Paged Pool Usage": process.QuotaNonPagedPoolUsage,
        "Session ID": process.SessionId,
        "Private Page Count": process.PrivatePageCount,
        "OS Name": process.OSName,
        "OS Creation Class Name": process.OSCreationClassName,
    }
    process_data.append(process_info)

# Create Dataframes
process_df = pd.DataFrame(process_data)

# Put dataframes into Excel
process_df.to_excel('running_processes.xlsx', index=True)
