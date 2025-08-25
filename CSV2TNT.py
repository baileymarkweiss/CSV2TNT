#Bailey Mark Weiss 26 January 2025
#   updated, cleaned, and annotated 24 August 2025

#Python code to translate google sheets to TNT file for matrices and Docx file for character statements
#Read Manual for explanation on layout of sheets
#Ensure that the Google Sheet is shared "Anyone with the link can edit"


#-----------------------Import packages-----------------------
import easygui
from tkinter import filedialog
from CSV2TNTFunctions import *
#---------------------------------------------------


#-----------------------Initialising variables-----------------------
UsePrev = None
FileName = ""
GoogleSheetsKEY = ""
CharactersGID = "None"
StatementsGID = "None"
Characters = False
Statements = False
#---------------------------------------------------


#-----------------------Choosing and making directory-----------------------
path = str(filedialog.askdirectory() +"/MatrixExport")
MakeFolder(path)
#---------------------------------------------------


#-----------------------loading variables-----------------------
UsePrev, FileName, GoogleSheetsKEY, CharactersGID, StatementsGID = PrevInfo(path) #loading data using tupple
#---------------------------------------------------


#-----------------------Asks user what to compile-----------------------
if  UsePrev == None: #Using previously found google sheet
    UsePrev = str(easygui.enterbox("Use previous Google Sheets saved as: '"+FileName+"'. This will redownload the matrix. (Y/N):"))
AskStatements = str(easygui.enterbox("Do you want to compile Statements? (Y/N):")) #compile Statements
AskCharacters = str(easygui.enterbox("Do you want to compile Characters? (Y/N):")) #compile Characters
#---------------------------------------------------


#-----------------------Error catching and loading booleans-----------------------
if AskStatements == "Y" or AskStatements == "y":
    if StatementsGID == "None" and (UsePrev == "Y" or UsePrev == "y"):
        print("No previous Statements GID for "+ FileName + " found!") #statments GID not found
        StatementsGID = str(easygui.enterbox("Previous GID not found!\nStatements Sheet GID:")) #input missing GID
    Statements = True #will run statment code
if AskCharacters == "Y" or AskCharacters == "y":
    if CharactersGID == "None" and (UsePrev == "Y" or UsePrev == "y"):
        print("No previous Character GID for "+ FileName + " found!") #characters GID not found
        CharactersGID = str(easygui.enterbox("Previous GID not found!\nCharacters Sheet GID:")) #input missing GID
    Characters = True #will run characters code
#---------------------------------------------------


#-----------------------New data input-----------------------
if UsePrev == "N" or UsePrev == "n":
    FileName = str(easygui.enterbox("File Name:")) #file name
    GoogleSheetsKEY = str(easygui.enterbox("Google Sheets Key:")) #Google Sheets key
    if Statements == True:
        StatementsGID = str(easygui.enterbox("Statements Sheet GID:")) #statements GID
    if Characters == True:
        CharactersGID = str(easygui.enterbox("Characters Sheet GID:")) #characters GID
#---------------------------------------------------


#-----------------------Writes data to InfoFile-----------------------
PrevInfoFile = open(path+"/InfoFile.txt", 'w')
PrevInfoFile.write("Name#"+FileName+
                "\nGoogleSheetsKEY#"+GoogleSheetsKEY+
                "\nCharactersGID#"+CharactersGID+
                "\nStatementsGID#"+StatementsGID)
PrevInfoFile.close()
#---------------------------------------------------


#-----------------------Run-----------------------
if Characters == True:
    createTNT(path, FileName, GoogleSheetsKEY, CharactersGID) #runs the function to create the tnt file
if Statements == True:
    CreateCharStatements(path, FileName, GoogleSheetsKEY, StatementsGID) #runs the function to create the docx file
#---------------------------------------------------