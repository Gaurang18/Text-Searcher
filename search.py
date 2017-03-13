#Import required modules
import fileinput
from shutil import move
from os.path import abspath, join, splitext, split
from os import mkdir, walk, remove
import win32com.client
import PyPDF2 as pyPdf

################
#Create lists to hold file names
file_list = list()
file_move_list = list()

#Define file extensions which need to be converted
excel_set = [".xls", ".xlsx", ".xlsm", ".xlsb"]
msword_set = [".doc", ".docx"]

################
#Define functions
def getFileList( searchdirectory ):
    #Get a list of all items in the directory to search
    for (dirpath, dirnames, filenames) in walk( searchdirectory ):
        for path in [ abspath( join( dirpath, filename ) ) for filename in filenames ]:
            file_list.append( path )

def searchFiles( readfilelist, movefilelist, searchstring ):
    #Get plain text from each file and search for searchstring
    for filename in readfilelist:
        ext = splitext( filename )[1]
        #Check filenames
        if searchstring in filename:
            movefilelist.append( filename )
            readfilelist.remove( filename )
        #Check if file is a pdf
        elif ext == ".pdf":
            content = getPDFContent( filename )
            if searchstring in content:
                movefilelist.append( filename )
        #Check if file is a word document
        elif ext in msword_set:
            app = win32com.client.Dispatch('Word.Application') 
            doc = app.Documents.Open( filename ) 
            if searchstring in doc.Content.Text:
                movefilelist.append( filename )
            app.Quit()
        #Check if file is an excel workbook/spreadsheet
        elif ext in excel_set:
            app = win32com.client.Dispatch( 'Excel.Application' )
            fileDir, fileName = split( filename )
            nameOnly = splitext( fileName )
            newName = nameOnly[0] + ".csv"
            outCSV = join( fileDir, newName )
            workbook = app.Workbooks.Open( filename )
            workbook.SaveAs(outCSV, FileFormat=24) # 24 is csv format
            workbook.Close(False)
            for line in open( outCSV, mode='r' ):
                if searchstring in line:
                    movefilelist.append( filename )
            app.Quit()
            remove( outCSV )
        #Assume all other files are plain text
        elif ext == ".txt":
            txtFile = open(filename, mode='r')
            for line in txtFile:
                if searchstring in line:
                    movefilelist.append( filename )
            txtFile.close()
        else:
            print(filename + " is not reconized")
        #readfilelist.remove( filename )

def moveFiles( movelist, destinationdirectory ):
    mkdir( destinationdirectory )
    for path in movelist:
        #Move the files to the destination folder
        move( path, destinationdirectory )

    print( 'Done' )

def getPDFContent( filename ):
    content = ""
    fd = file(filename, 'rb')
    pdf = pyPdf.PdfFileReader( fd )
    # Extract text from each page and add to content
    for i in range( 0, pdf.getNumPages() ):
        content += pdf.getPage(i).extractText() + " \n"
    fd.close()
    return content

################
#Run as main
if __name__=='__main__':
    search_directory = input( 'Enter the path of the directory you wish to search through, in this format "C:\Users\admin\folder" : ' )
    search_string = input( 'Enter the search term in quotes: ' )
    destination_directory = input( 'Enter the name of the new directory which will contain the moved files, in this format"C:\Users\admin\folder" : ' )

    getFileList( search_directory )
    searchFiles( file_list, file_move_list, search_string )
    moveFiles( file_move_list, destination_directory )