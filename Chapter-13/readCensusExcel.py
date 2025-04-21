import os
import census2010

pilipinas = census2010.allData['PH']['Philippines']['pop']
print('The population of the Philippines is ' + str(pilipinas))