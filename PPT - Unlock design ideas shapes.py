import logging, os, sys, zipfile
from tkinter.filedialog import askopenfilename
from lxml import objectify, etree
"""Unlocks the shapes created by Design Ideas  in PowerPoint.
For this, it unzips the pptx. Then searches for Shape Locks tags (a:spLocks), and modifies their attributes.
It also removes the tags adec:decorative and p16:designElem belonging to the same shape.
The result is saved in a file suffixed with _mod.pptx"""

CURRENT_DIR = os.path.dirname(sys.argv[0]) #Path of current script
#List of all attributes that may be locked and we want to unlock
LOCK_ATTR = ("noGrp", "noRot", "noChangeAspect", "noMove", "noResize", "noEditPoints", "noAdjustHandles", "noChangeArrowheads", "noChangeShapeType", "noTextEdit")
DECORS_XPATH = ( "../..//adec:decorative", "../..//p16:designElem") #XPATH of other tags that must be removed
LOGGING_LEVEL = logging.INFO #DEBUG, INFO, WARNING, ERROR, CRITICAL

def choose_file(mydir=CURRENT_DIR, mytitle='Choose a ppt'):
    return askopenfilename(initialdir=mydir, title=mytitle, filetypes=[("Powerpoint Files", "*.pptx")])

def unzip_file(zip_name, unzip = True): #Unzip File
    unzip_dir = os.path.splitext(zip_name)[0]
    with zipfile.ZipFile(zip_name,"r") as zip_ref:
        if unzip: zip_ref.extractall(unzip_dir)
    return unzip_dir

def zip_directory(zip_dir, zip_name, compression = zipfile.ZIP_DEFLATED):
    with zipfile.ZipFile(zip_name,"w", compression) as zip_ref:
        for root, dirs, files in os.walk(zip_dir):
            root_rel = os.path.relpath(root, zip_dir) #keep only relative path
            zip_ref.write(root, root_rel) #Write dir with relative path (needed for empty dirs)
            for file in files:
                filename = os.path.join(root, file)
                if os.path.isfile(filename): # regular files only
                    zip_ref.write(filename, os.path.join(root_rel, file)) #Write file with relative path

def main():
    ppt_path = choose_file(CURRENT_DIR)
    logging.basicConfig(level=LOGGING_LEVEL, format="{asctime} {message}", style="{", datefmt="%H:%M:%S")
    if ppt_path:
        unzip_dir = unzip_file(ppt_path, True) #unzip .pptx
        ppt_dir = os.path.join(unzip_dir, "ppt", "slides") #got to dir containing xml of slides
        _, _, files = next(os.walk(ppt_dir))
        for file in files: #loop on files of slides
            xml_slide, ext = os.path.splitext(file)
            if ext.upper()==".XML": #check this is a xml file
                logging.debug(xml_slide)
                with open(os.path.join(ppt_dir, file), 'rb') as xml_reader: #open file
                    xml_data = objectify.parse(xml_reader) #parse XML to ElementTree
                    root = xml_data.getroot()   #get root Element
                for elem in root.findall(".//a:spLocks", namespaces=root.nsmap):
                    found = False
                    if len(elem.attrib) > 1: #Sometimes a tag has only 1 attribute noGrp. If it is the case, we leave it
                        for attr in elem.attrib.keys():
                            if attr in LOCK_ATTR:
                                found = True
                                elem.attrib.pop(attr) #Remove lock attributes
                    if found:
                        nsmap_all = {}
                        for ns in root.xpath('//namespace::*'): #Also looks at namespaces that are not at root. Ex: adec, a16
                            if ns[0]: # Removes the None namespace, neither needed nor supported.
                                nsmap_all[ns[0]] = ns[1]
                        for decor in DECORS_XPATH:
                            decors = elem.findall(decor, namespaces=nsmap_all)
                            if len(decors) == 1: #Check if tag was found
                                decors[0].getparent().remove(decors[0]) #Remove tags in DECORS
                with open(os.path.join(ppt_dir, file), 'wb') as xml_writer: #overwrite file with modified root
                    xml_writer.write(etree.tostring(root))
        zip_directory(unzip_dir, unzip_dir + "_mod.pptx") #Zip directory back to pptx

if __name__ == '__main__':
    main()