import os
import time

#os.chdir('C:/Users/USER/OneDrive - 경희대학교/오버워치 승패 기록표')
#os.system('poetry env activate')

os.system('poetry run source_generate_report_20250822.py')

while True :
    userinput = input('\nPress enter to exit')
    if userinput == '' :
        time.sleep(0.5)
        break
