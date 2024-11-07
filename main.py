import sys, getopt
import DataEngine

def main(argv):
    cik = ''
    print('I am hereI am hereI am hereI am hereI am here')
    try:
        opts, args = getopt.getopt(argv, "fhc:")
    except getopt.GetoptError:
        print ('app.py -c <cik> [-f]')
        sys.exit(2)

    forceDownload = False
    for opt, arg in opts:
        if opt == '-h':
            print ('app.py -c <cik>')
            sys.exit()
        if opt  == '-f':
            forceDownload = True
        elif opt in ('-c'):
            print('I am here')
            cik = arg            
    print ('CIK "', cik, forceDownload)    
    DataEngine.save('0000883241', False)    
    #DataEngine.save(cik, forceDownload)

if __name__ == "__main__":
   main(sys.argv[1:])

