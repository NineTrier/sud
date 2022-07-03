import subprocess
import re
from time import sleep, time
import threading

tasks = subprocess.check_output(['tasklist']).split(b'\r\n')
for i in tasks:
    print(i.decode(errors="ignore")) if i.decode(errors="ignore").find('W') != -1 else ""
