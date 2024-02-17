import os
import re
import tempfile
import PySimpleGUI as sg
import sys
import shutil
import PIL.Image
import io
import base64

############################################################
# Error Class
############################################################

class SpaceInFilepathError(Exception):
    """
    Spaces are not allowed when specifying file-paths.
    """
    pass

class NoFilepathError(Exception):
    """
    No filepath has been specified.
    """
    pass

class SpaceInFileNameError(Exception):
    """
    Spaces are not allowed in filenames (Not to be confused with the file title, where spaces are allowed).
    """
    pass

class NoFilenameError(Exception):
    """
    No filename has been specified.
    """
    pass

class NoOutTypeError(Exception):
    """
    Please select the output file type.
    """
    pass

class NoInTypeError(Exception):
    """
    Please select the input file type.
    """
    pass

class NoProductSelectedError(Exception):
    """
    Please select a product/template type.
    """
    pass

class NoHeaderSelectedError(Exception):
    """
    Please select the appropriate header.
    """
    pass

class NoFileTitleError(Exception):
    """
    Please enter a file title.
    """
    pass

class NoTitlePageTemplateError(Exception):
    """
    Please select a title page template.
    """
    pass

############################################################
# GUI Configurations
############################################################

CompanyTheme = {'BACKGROUND': '#404040',
                'TEXT': 'white',
                'INPUT': 'white',
                'TEXT_INPUT': '#000000',
                'SCROLL': '#404040',
                'BUTTON': ('#404040', '#61C9A8'),
                'PROGRESS': ('#404040', '#61C9A8'),
                'BORDER': 0,
                'SLIDER_DEPTH': 0,
                'PROGRESS_DEPTH': 0}

sg.theme_add_new('Company-Colours', CompanyTheme)
sg.theme('Company-Colours')

productList = ['General', 'Product 1', 'Product 2', 'Product 3'] #list of different product templates.
headerList = [' ','CONFIDENTIAL INFORMATION', 'HIGHLY CONFIDENTIAL INFORMATION', 'RESTRICTED INFORMATION', 'INTERNAL ONLY'] #List of different headers.
version = '3.5'
houseFont = 'Arial'

#sg.ChangeLookAndFeel('GreenTan')
form = sg.FlexForm(f'Eitri {version}', default_element_size=(50, 1), no_titlebar=False, alpha_channel=1, grab_anywhere=True, resizable=True)
left_col = [
    [sg.Text(f'EITRI', size=(30, 1), font=('Arial', 25), text_color='#61C9A8')],
    [sg.Text('Offline Markdown Converter. \nFor more information, please visit https://github.com/s-junlee98/Eitri.', font=(houseFont, 13), tooltip='Eitri is also the dwarf who created Thanos\' Infinity Gauntlet and Thor\'s Stormbreaker.')],
    [sg.Text('', font=(houseFont, 13))],
    [sg.Text('Choose a folder containing files to be compiled. \nPlease ensure there are no spaces in the directory or the file names.',
     tooltip='The directory cannot have spaces.', font=(houseFont, 13))],
    [sg.Text('Selected Folder:', font=(houseFont, 13), auto_size_text=True, justification='right'),
     sg.InputText(''), sg.FolderBrowse()],
    [sg.Text('\nChoose the disclaimer PDF. Please ensure there are no spaces.', tooltip='Visit docs.Company.com for the Confidentiality Guide.', font=(houseFont, 13))],
    [sg.Text('Selected File:', font=(houseFont, 13), auto_size_text=True, justification='right'),
     sg.InputText(''), sg.FileBrowse(initial_folder="G:\Development\Confidentiality-Guide")],
    [sg.Text('\nConverting from:', font=(houseFont, 13))],
    [sg.Listbox(values=('.md', '.docx'), no_scrollbar=True, tooltip='Input file format', font=(houseFont, 13), size=(15, 2))], #the blank option is because the listbox wants a list with 2+ elements
    [sg.Text('Converting to:', font=(houseFont, 13))],
    [sg.Listbox(values=('.pdf', '.tex'), no_scrollbar=True, tooltip='PDF or LaTeX', font=(houseFont, 13), size=(15, 2))], #out type
    [sg.Text('Select your document template:', font=(houseFont, 13))],
    [sg.Listbox(values=productList, tooltip='General for product agnostic documents.', no_scrollbar=True, font=(houseFont, 13), size=(32, 4))], #product type
    [sg.Text('Select your confidentiality header:', font=(houseFont, 13))],
    [sg.Listbox(values=headerList, tooltip='Each option corresponds to a confidentiality category.', no_scrollbar=True, font=(houseFont, 13), size=(32, 5))], #product type
    [sg.Text('Enter the filename.', font=(houseFont, 13))],
    [sg.InputText('')],
    [sg.Text('Enter the title to be displayed on the front page:', tooltip='You can use \\linebreak to start a new line.', font=(houseFont, 13))],
    [sg.InputText('')],
    [sg.Text('Enter any relevant version number:', tooltip='Leave blank if not applicable.', font=(houseFont, 13))],
    [sg.InputText('')],
    [sg.Text('Select the title page design:', font=(houseFont, 13))],
    [sg.InputText('', enable_events=True, key='-SELECT_DESIGN-'), sg.FileBrowse(initial_folder='G:\Development\Eitri\TitlePageTemplates')],
    [sg.Button('GENERATE DOCUMENT', size=(20,2), expand_x=True)],
    [sg.Text('\ns.junlee98@gmail.com', size=(59,2), justification='center', font=(houseFont, 11))]
     ]

right_col = [
    [sg.Text('Selected Title Page Design:', font=(houseFont, 13))],
    [sg.Text(size=(40,1), key='-TOUT-')],
    [sg.Image(key='-IMAGE-')],
]

layout = [[sg.Column(left_col, scrollable=True, vertical_scroll_only=True, expand_x=True, expand_y=True), sg.VSeparator(), sg.Column(right_col, element_justification='c')]]

############################################################
# Section-Object Class
############################################################

class Section:
    def __init__(self, section_str):
        self.Title = Section.getSecTitle(section_str)
        self.Level = Section.getSecTitleLv(section_str)
        self.IsGreen = Section.hasGreens(section_str)
        self.Content = Section.getGreens(section_str)

    @staticmethod
    def getSecTitle(section_str): #returns any headers which are tagged in green
        '''
        Gets the section title of the input text
        Input: One section-text, already split (str)
        Output: The title beginning with '#' (str)
        '''
        
        secTitlePat = r"(^[#]{1,6}\s[^\n]*)" #catches headings beginning with #s
        sectionTitle = re.findall(secTitlePat, section_str)
        return sectionTitle

    @staticmethod
    def getSecTitleLv(section_str): #function that fetches the level of the header (h1, h2, h3, etc...)
        '''
        Retrieves the header level
        Input: One text-section from .md (string)
        Output: Header level (# = level 1) (int)
        '''
        hashpat= "([#]{1,6}[\n]*)" #regex for filtering for #s
        hashes = re.findall(hashpat, section_str) #isolates #s. Is a list.
        level = len(hashes[0]) #counts the no. of #s in index 0. There should only ever be index 0
        return level

    @staticmethod
    def hasGreens (text):
        '''
        Scans through any bit of string to see if it has <green>s
        Input: file (file)
        Output: Boolean.
        '''
        greenpat = r"<green>"
        if re.search(greenpat, text):
            return True
        else:
            return False

    @staticmethod
    def getGreens(section_str): 
        '''
        Searches through one file, gets all sections marked with <green>
        Input: file-text (str)
        Output: text which are marked with <green>
        '''
        greenSecPat = r"((^[#]*\s<green>[^<]*<\/green>)[^#]*)|(.*(?=<green>)(<green>)[^\n]*)"
        greenSections = re.findall(greenSecPat, section_str)
        return greenSections
    
############################################################
# General Functions Block
############################################################

def convert_to_bytes(file_or_bytes, resize=None): #taken from PySimpleGUI cookbook. 
    '''
    Will convert into bytes and optionally resize an image that is a file or a base64 bytes object.
    Turns into  PNG format in the process so that can be displayed by tkinter
    :param file_or_bytes: either a string filename or a bytes base64 image object
    :type file_or_bytes:  (Union[str, bytes])
    :param resize:  optional new size
    :type resize: (Tuple[int, int] or None)
    :return: (bytes) a byte-string object
    :rtype: (bytes)
    '''
    if isinstance(file_or_bytes, str):
        img = PIL.Image.open(file_or_bytes)
    else:
        try:
            img = PIL.Image.open(io.BytesIO(base64.b64decode(file_or_bytes)))
        except Exception as e:
            dataBytesIO = io.BytesIO(file_or_bytes)
            img = PIL.Image.open(dataBytesIO)

    cur_width, cur_height = img.size
    if resize:
        new_width, new_height = resize
        scale = min(new_height/cur_height, new_width/cur_width)
        img = img.resize((int(cur_width*scale), int(cur_height*scale)), PIL.Image.Resampling.LANCZOS)
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    del img
    return bio.getvalue()

def resource_path(relative_path): #https://stackoverflow.com/questions/7674790/bundling-data-files-with-pyinstaller-onefile/44352931#44352931
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def getfiles(directory, ext):# gets the files from the directory
    '''
    Fetches all files of the specified extension.
    Input: Folder directory "directory" (str), file extension "ext" (str)
    Output: A list of files with the same extension (list)
    '''
    files = []
    for (dirpath, dirnames, filenames) in os.walk(directory):
        for x in filenames:
            if x.endswith(ext):
                files.append(os.path.join(dirpath, x))
    return files

def getInType():
    print ("What filetype would you like to convert FROM? E.g: .md ")
    return input()

def getoutput(): #gets the name of the output file
    print("What is the name of your output file? (include file extension: .pdf or .tex)")
    return input()

def getouttype(output): #automatically fetches the output file type
    return os.path.splitext(output)[1]

def getversion(): #gets the version (str) number
    print("What version of the product is this for? Just leave blank if N/A.")
    return input()

def mktempdir(folder): #creates a temp directory 
    direc = f"{directory}\\{folder}\\"
    try:
        os.mkdir(direc)
    except OSError:
        print("tempdir creation failed (already exists)")
    return direc

def getTempImage(res_dir):
    '''
    Copy the images from the application temporarily into where the md files are
    Input: Resource directory within the application
    Output: Action - copies temp images into active directory.
    '''
    imgdir = f"{res_dir}\\img"
    targetdir = f"{directory}\\tempimg"
    imgext='.png'
    try:
        os.mkdir(targetdir)
        
    except OSError:
        print("tempdir creation failed (already exists)")

    for (dirpath, dirnames, filenames) in os.walk(imgdir):
        for x in filenames:
            if x.endswith(imgext):
                shutil.copy(f"{imgdir}\\{x}", targetdir)

    return targetdir

def removeTempImage():
    '''
    Removes temp images - hardcoded names!!! May need improvements to the code.
    '''
    for item in f"{directory}\\tempimg":
        os.remove(item)
    return

def gettemplate(product): #gets template if output file is .pdf or .tex.
    '''
    Returns the location of the 4 templates (non-client)
    Input: None
    Output: List with 4 items (list)
    '''
    templateList = []
    template=''
    if outtype == ".pdf" or outtype == ".tex":
        templateList.append(f"--template={res_dir}\\general.tex") 
        templateList.append(f"--template={res_dir}\\product_1.tex") 
        templateList.append(f"--template={res_dir}\\product_2.tex") 
        templateList.append(f"--template={res_dir}\\product_3.tex") 

        if product == 'General':
            template = templateList[0]
        elif product == 'Product~1':
            template = templateList[1]
        elif product == 'Product~2':
            template = templateList[2]
        elif product == 'Product~3':
            template = templateList[3]

    return template

def greenH2(filetext): #sees if the H2 header (should be largest header in file) is green
    '''
    Checks the h2 header to see if it's wrapped in <green>s.
    Input: All of the text from one .md file (str)
    Output: Boolean
    '''
    greenh2Pat = '##\s[^#]*(?=\\n\\n)'
    h2Header = re.findall(greenh2Pat, filetext)
    if not h2Header:
        return False
    elif re.search('<green>', h2Header[0]):
        return True
    else:
        return False

def gettitle (): #changes spaces into ~s for pandoc conversion.
    print("What is the title of the document?")
    title = blankInputToSpace(input())
    return title

def blankInputToSpace(text):
    """
    Converts a blank input (when no string is provided) into single tilde for latex formatting.
    Input: string.
    Output: string with either ~ if no input, or all spaces converted to ~ if string exists.
    """
    if text == '':
        return '~'
    else:
        return re.sub(r"\s", r"~", text)

def spaceToDash(text):
    """
    Converts a space to dash "-".
    Input: Any string.
    Output: The same string, but spaces replaced with dashes.
    """
    return re.sub(r"\s", r"-", text)

############################################################
# Client-Facing Doc Generation
############################################################

def getdata(file): #gets the whole text as string
    with open(file, encoding="utf8") as f:
        filetext = f.read()
    return filetext

def getheaders(filetext): #returns any headers which are tagged in green
    '''
    Returns a list of headers including the green tags
    Input: Text in the file being scanned (str)
    Output: List of headers (list)
    '''
    headerpattern = re.compile(r"(^[#]*\s<green>[^<]*<\/green>)") #the regex catches just the "<green>text</green>""
    headers = []
    for match in re.findall(headerpattern, filetext):
        headers.append(match)
    return headers

def removeRed (file, tempdirec):
    '''
    Takes away anything that's marked in <red></red> tags
    Input: The directory of a .md file (str)
    Output: The input file but with red tags removed. Saves the file in the client directory
    '''
    filetext = getdata(file)
    redPat = r'((^[#]*\s<red>[^<]*<\/red>)[^#]*)|(.*(?=<red>)(<red>)[^\n]*)'
    h1Pat = r'# Pipeline'
    newText = re.sub(h1Pat, ' ', filetext)
    newText = re.sub(redPat, '', newText)
    with open(f'{tempdirec}{os.path.basename(file)}','w', encoding="utf8") as MD:
        MD.write(newText)
    return

############################################################
# Green Tag Removal
############################################################

def clientregex(origmd): #swaps the tags so that it just looks normal for client-facing docs
    clientfiles = []
    pats = {}

    pats[(re.compile(r"(\<green\>)", re.IGNORECASE))] = r""
    pats[(re.compile(r"(\<\/green\>)", re.IGNORECASE))] = r""
    pats[(re.compile(r"(\<red\>)", re.IGNORECASE))] = r"\\textcolor{red}{"
    pats[(re.compile(r"(\<\/red\>)", re.IGNORECASE))] = r"}"

    for file in origmd:
        filetext = getdata(file)
        if greenH2(filetext): # if a file has green H2
            clientfiles.append(replacetags(file, pats, True))

        else: # if there are no green tags in the whole file
           # origmd.remove(file)
           pass
    
    # for file in clientfiles:
    #     removeRed(file)
    
    return clientfiles

def mainregex(origmd): #replaces <green> tags with LaTeX formatting for font colour.
    newfiles = []
    pats = {}

    pats[(re.compile(r"(\<green\>)", re.IGNORECASE))] = r"\\textcolor{green}{"
    pats[(re.compile(r"(\<\/green\>)", re.IGNORECASE))] = r"}"
    pats[(re.compile(r"(\<red\>)", re.IGNORECASE))] = r"\\textcolor{red}{"
    pats[(re.compile(r"(\<\/red\>)", re.IGNORECASE))] = r"}"

    for file in origmd:
        newfiles.append(replacetags(file, pats, False))
    return newfiles

############################################################
# PDF Generation
############################################################

def replacetags(file,patterns,isclient):
    '''
    Replaces <green> and </green> tags with LaTeX formatting
    Input: file (str), patterns (dict), isclient (boolean)
    Output: replaces tags on the files (none)
    '''
    #point to correct temp dir
    tempdirec='' #initialising temp directory for use outside the if statements
    data = '' #initialising data for use outside the if statements
    if(isclient):
        tempdirec = clienttempdir
        removeRed(file, tempdirec)
        data = getdata(f"{tempdirec}{os.path.basename(file)}")

    else:
        tempdirec = tempdir
        data = getdata(file)
    
    #read contents orig md file, apply replacements
    

    for pat in patterns.keys():
        data = re.sub(pat, patterns[pat], data)
    
    #get temporary file
    temp = f"{tempdirec}{os.path.basename(file)}"

    #save temporary files
    t = open(temp, "w", encoding="utf8")
    t.write(data)
    t.close()
    
    return temp   

def pandocMarkdownIn(inputFiles, disclaimerDirectory, isclient): #passes text to a cmd to use pandoc conversion
    out = output
    tempdir = ''
    if (isclient):
        tempdir = clienttempdir
        out = "External-" + out
        #templatetype = template[1] #Second item in the list has the default_client.tex directory
    else:
        tempdir = tempdir
        #templatetype = template[0] #First item in the list has the default.tex directory 

    #this line runs regardless (at the moment)
    pandoccommand = f"cd {directory} && pandoc {inputFiles} -s -o {directory}\\{out} {template} --pdf-engine=xelatex --top-level-division=part --resource-path={directory} --variable=header:{header} --variable=titledesign:{titlePage} --variable=versionnumber:{version} --variable=title:{docTitle} --variable=disclaimerPDF:{disclaimerDirectory}" 
    
    #purely printing for comms, does nothing else
    if (isclient):
        print("Generating your client-facing .pdf...")
    else:
        print("Generating your internal-facing .pdf...")
    
    try:
        if inputFiles == '': #checks to see if there are any input files in the specified directory
            print (f'Friendly Information: No temporary files in {tempdir}.\nThis message may indicate that only the internal-version has been generated.')
        else:
            os.system(pandoccommand)
    except:
        print("Error while running Pandoc!")

def pandocDocxIn(inputFiles):
    out = '.md'
    tempdir = ''
    pandoccommand = f"cd {directory} && pandoc --extract-media ./myMediaFolder {inputFiles} -o {directory}\\tempMD{out}"

    try:
        if inputFiles == '':
            print (f'Information: No temporary files in {tempdir}.\nThis message may indicate that only the internal-version has been generated.')
        else:
            os.system(pandoccommand)
    except:
        print("Error while running Pandoc (error while converting .docx to .md)!")

    return

def killtemp():
    shutil.rmtree(tempdir)
    shutil.rmtree(clienttempdir)
    shutil.rmtree(tempimgdir)

############################################################
# Functions for the Main Method
############################################################

def markdownToOutTypeMethod (directory, disclaimerDirectory, intype):

    origmdfiles = getfiles(directory, intype) #get raw md files
    mainfiles = mainregex(origmdfiles) #do some regex replacements to get the internal-facing and external-facing md files
    clientfiles = clientregex(origmdfiles)
    tempfiles = getfiles(tempdir, intype) #get temp files
    clienttempfiles = getfiles(clienttempdir, intype)
    tempimg = getfiles(f"{directory}\\tempimg", '.png')
    mainmd = ' '.join(tempfiles) #get markdown
    clientmd = ' '.join(clienttempfiles)

    pandocMarkdownIn(mainmd, disclaimerDirectory, False) #send md to pandoc
    pandocMarkdownIn(clientmd, disclaimerDirectory, True)

    killtemp() #delete temp files
    #removeTempImage()

    print("Done! Program has finished running.") #hail the arrival of a new age
    return

def docxToMarkdownMethod (directory,intype):
    origdocxfiles = getfiles(directory,intype)
    tempimg = getfiles(f"{directory}\\tempimg", '.png')
    docxListString = ' '.join(origdocxfiles)

    pandocDocxIn(docxListString)

    print("Done! Program has finished running.")
    return

############################################################
# Reading the User Inputs from GUI
############################################################

EitriGUI = form.Layout(layout) #displays the GUI. Doesn't close until program is exited.

while True:
    event, userInput = EitriGUI.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    if event == sg.WIN_CLOSED or event == 'Exit':
        break
    if event == '-SELECT_DESIGN-':
        try:
            filename = os.path.join(userInput['-SELECT_DESIGN-'])
            EitriGUI['-TOUT-'].update(filename)
            size=(248, 350)
            EitriGUI['-IMAGE-'].update(data=convert_to_bytes(filename, size))
        except Exception as E:
            print(f'** Error {E} **')
            pass        # something weird happened making the full filename
    if event == 'GENERATE DOCUMENT':
        try: 
            directory = userInput[0] #directory of the content
            disclaimerDirectory = userInput[1] #file location of the disclaimer PDF
            intypeList = userInput[2] #input type (.md)
            outtypeList = userInput[3] #output type (.pdf or .tex)
            productList = userInput[4] #product name (needs to match with the .tex templates)
            headerList = userInput[5] #header selected (blank, confidential, highly confidential)
            filename = userInput[6] #file name (section before file extension)
            docTitle = userInput[7] #document title (text, no formatting requirements)
            version = userInput[8] #version (simple text, not required)
            titlePageTemplate = userInput['-SELECT_DESIGN-'] #title page template

            if ' ' in directory:
                raise SpaceInFilepathError
            if directory =='':
                raise NoFilepathError
            if ' ' in disclaimerDirectory:
                raise SpaceInFilepathError
            if disclaimerDirectory == '':
                raise NoFilepathError
            # if ' ' in filename:
            #     raise SpaceInFileNameError
            if filename =='':
                raise NoFilenameError
            if len(intypeList)==0:
                raise NoInTypeError
            if len(outtypeList)==0:
                raise NoOutTypeError
            if len(productList)==0:
                raise NoProductSelectedError
            if len(headerList)==0:
                raise NoHeaderSelectedError
            if docTitle=='':
                raise NoFileTitleError
            if titlePageTemplate == '':
                raise NoTitlePageTemplateError
            
            #print(input_values)
            filenameDetails = [spaceToDash(filename), outtypeList[0]] #makes it compatible with existing code.
            directory = str(directory)
            disclaimerDirectory = str(disclaimerDirectory)
            intype = str(intypeList[0])
            output = ''.join(filenameDetails)
            outtype = str(outtypeList[0])
            product = blankInputToSpace(productList[0])
            header = blankInputToSpace(headerList[0])
            version = blankInputToSpace(version) #get version of the product
            docTitle = blankInputToSpace(docTitle) #swaps space with ~ to make pandoc happy
            titlePage = str(titlePageTemplate)

        except NoFilepathError:
            sg.popup_error('Please choose a filepath containing the input files.')
            continue
        except SpaceInFilepathError:
            sg.popup_error('No spaces allowed in the filepath!')
            continue
        except NoFilenameError:
            sg.popup_error('Please enter a filename.')
            continue
        except SpaceInFileNameError:
            sg.popup_error('No spaces allowed in the file name! Not to be confused with file title, where spaces are allowed.')
            continue
        except NoInTypeError:
            sg.popup_error('Please select the type of the input file.')
            continue
        except NoOutTypeError:
            sg.popup_error('Please select the type of the output file.')
            continue
        except NoProductSelectedError:
            sg.popup_error('Please select the product/template!')
            continue
        except NoHeaderSelectedError:
            sg.popup_error('Please select a header!')
            continue
        except NoFileTitleError:
            sg.popup_error('Please enter a file title!')
            continue
        except NoTitlePageTemplateError:
            sg.popup_error('Please select a title page template!')
        else:
            res_dir = resource_path('res')
            #res_dir = directory #turn this flag on when testing.
            tempdir = mktempdir("temp") #create temp directories
            clienttempdir = mktempdir("clienttemp")
            tempimgdir = getTempImage(res_dir)
            template = gettemplate(product) #get template

            if intype == '.md': #if the input is a markdown
                markdownToOutTypeMethod(directory, disclaimerDirectory, intype)
                sg.popup('Done! Markdown file(s) have been converted.')
                continue

            elif intype == '.docx': #if the input is docx
                docxToMarkdownMethod(directory, intype)
                markdownToOutTypeMethod(directory, disclaimerDirectory, '.md')
                sg.popup('Done! Docx file(s) have been converted.')
                continue

EitriGUI.close()