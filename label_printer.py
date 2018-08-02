import time
import win32com
import win32com.client

labels_to_print = """
Malaga patching/cabling
""".splitlines()

labels_to_print = [x for x in labels_to_print if len(x) > 0]

shell = win32com.client.Dispatch("WScript.Shell")
# shell.Run("Dymo Label Light")
# time.sleep(1)
shell.AppActivate('DYMO Label Light')
for label in labels_to_print:
    shell.SendKeys("^n")
    shell.SendKeys(label)
    shell.SendKeys("^p")
    time.sleep(5)