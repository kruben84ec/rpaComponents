import subprocess
import sys
try:
    if(len(sys.argv)>1):
        print('Giskard say:','Limpiando los temporales')
        limpiador_temporal = str(sys.argv[1])
        subprocess.call([limpiador_temporal])
        print('Giskard say:','Mantando los excel')
        kill_excel = str(sys.argv[2])
        subprocess.call([kill_excel])
except IOError as error:
    print('Giskard say:',str(error))
