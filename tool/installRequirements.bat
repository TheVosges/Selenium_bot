if not defined in_subprocess (cmd /k set in_subprocess=y ^& %0 %*) & exit )
g:
cd G:\Python3.8\Portable Python-3.8.6 x64\App\Python
python -m pip install -r "I:\ICS\IC\Automation\16. Reconciliation Tool\tool\requirements.txt"