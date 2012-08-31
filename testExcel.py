import excelGeneral
import sys

def main(argv):

    #excelGeneral.checkSourceFormat("OTD Jul'12.xlsx")
    #excelGeneral.processExcel("OTD Jul'12.xlsx")

    excelGeneral.checkSourceFormat(argv)
    excelGeneral.processExcel(argv)

if __name__ == "__main__":
    main(sys.argv[1])
