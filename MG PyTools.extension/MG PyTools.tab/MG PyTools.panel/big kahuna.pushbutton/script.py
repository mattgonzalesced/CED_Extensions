import subprocess
from pyrevit import revit, script

# Initialize output manager
output = script.get_output()
output.close_others()

# Log start of execution
log_file = r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\script_execution_log.txt"

# Open the log file for the entire script execution
with open(log_file, "w") as log:
    log.write("Starting sequential execution of scripts...\n")

    # Step 1: Run the first IronPython script
    try:
        log.write("Running first script (IronPython)...\n")
        execfile(r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\initial filter.py")
        log.write("First script completed.\n")
    except Exception as e:
        log.write("Error running first script: {0}\n".format(str(e)))

    # Step 2: Run the second Python script
    try:
        log.write("Running second script (Python)...\n")
        subprocess.call(
            ["python", r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\EXT pandas.py"]
        )
        log.write("Second script completed.\n")
    except Exception as e:
        log.write("Error running second script: {0}\n".format(str(e)))

    # Step 3: Run the third IronPython script
    try:
        log.write("Running third script (IronPython)...\n")
        execfile(r"C:\Users\m.gonzales\OneDrive - CoolSys Inc\Desktop\CED\DevExtension\MG PyTools.extension\MG PyTools.tab\MG PyTools.panel\big kahuna.pushbutton\CSV selector output.py")
        log.write("Third script completed.\n")
    except Exception as e:
        log.write("Error running third script: {0}\n".format(str(e)))

    log.write("All scripts executed successfully.\n")

print("Execution complete. Check log file for details.")
