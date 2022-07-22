src = r"I:\ICS\IC\Service Now Reconciliation\temp\Dummy Files\\"
dst = r"I:\ICS\IC\Service Now Reconciliation\temp\Generation\\"
import os, shutil
shutil.copytree(src, dst, dirs_exist_ok=True)