import shutil
import subprocess
import sys
import os
from make_log import log_exceptions

try:
    shutil.rmtree('temp_files', ignore_errors=True)
    os.mkdir('temp_files')
    insname, fpath = sys.argv[1], sys.argv[2]
    subprocess.run(["python", 'pdf_' + insname + ".py", fpath])
except:
    log_exceptions()
pass
