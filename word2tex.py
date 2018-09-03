from docx import Document
from docx.shared import Inches
import re
import argparse
import zipfile
import base64
from pprint import pprint
import  xml.etree.ElementTree as ET
from re import sub
import os

#Bachelorarbeitverkürzt_Publikation2
image_insert_fullwidth = """\\begin{figure*}[ht]\centering
\includegraphics[width=0.9\linewidth]{FILENAME}
\caption{CAPTION_HERE}
\label{fig:FILENAME}
\end{figure*}CAPTION_FOOTNOTES_HERE"""
image_insert_column = """\\begin{figure}[ht]\centering
\includegraphics[scale=0.5]{FILENAME}
\caption{CAPTION_HERE}
\label{fig:FILENAME}
\end{figure}CAPTION_FOOTNOTES_HERE"""

parser = argparse.ArgumentParser(description='open sourcefile ', argument_default=argparse.SUPPRESS)
parser.add_argument('source', help='input file name')
parser.add_argument('--bibfile', help='input Bib file name', default="")
parser.add_argument('--pre', help='optional pre file name', default='huplc_pre.tex')
parser.add_argument('--post', help='optional post file name', default='huplc_post.tex')
# args = parser.parse_args()
args = parser.parse_args()

sourcefile = args.source
prefile = args.pre
postfile = args.post




# variable dumper for debug purposes
def dump(obj):
   for attr in dir(obj):
       if hasattr( obj, attr ):
           print( "obj.%s = %s" % (attr, getattr(obj, attr)))

# write tex tags by language code
def lang_switch(code, str):
    if re.match(r"ar-.*", code):
        return  "\\foreignlanguage{arabic}""{"+str+"}"
    if re.match(r"(zh|jp|kr)-.*", code):
        return "\\begin{CJK}{UTF8}{gbsn}"+str+"\\end{CJK}"
    return {
        'el-GR': "\\foreignlanguage{greek}""{"+str+"}",
        'he-IL': "\\foreignlanguage{hebrew}""{"+str+"}",
        'ru-RU': "\\foreignlanguage{russian}""{"+str+"}",
        # 'ro-RO': "\\foreignlanguage{romanian}""{"+str+"}",
        # 'sk-SK': "\\foreignlanguage{slovak}""{"+str+"}",
        # 'sv-SE': "\\foreignlanguage{swedish}""{"+str+"}",
        # 'tr-TR': "\\foreignlanguage{turkish}""{"+str+"}",
        # 'uk-AU': "\\foreignlanguage{ukrainian}""{"+str+"}",
    }.get(code, str)

# process runs and generate tex code
def process_runs(runs, in_caption = False):
    out = ""
    out2 = ""
    citavimode = False
    for index, r in enumerate(runs):
        is_element = str(r.__class__) == "<class 'xml.etree.ElementTree.Element'>"
        processed = ""

        try:
            if is_element:
                fldChar = r.findall('./{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldChar')
                if len(fldChar):
                    if fldChar[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType') == 'begin':
                        citavimode = True
                    if fldChar[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType') == 'end':
                        citavimode = False
            else:
                fldChar = r.element.xpath('./w:fldChar')
                if len(fldChar):
                    if fldChar[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType') == 'begin':
                        citavimode = True
                    if fldChar[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType') == 'end':
                        citavimode = False
        except:
            print('', end="")

        if not citavimode:
            # behandle den text
            if (is_element):
                for child in r.findall('./*'):
                    if (child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t' and child.text):
                        processed += child.text
                    if (child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab'):
                        processed += '\t'
            else:
                processed = str(r.text)
            processed = re.sub(r"(_|%|&|\\)",  r"\\\1", processed)
            processed = re.sub(r"α",  r"\\textalpha", processed)
            processed = re.sub(r"ʒ",  r"\\t{{Z}}", processed)
            processed = re.sub(r"\t",r"\\quad ",processed)
            if is_element:
                lang_element = r.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lang')
            else:
                lang_element = r.element.xpath('w:rPr/w:lang')
            if len(lang_element):
                langcode = 'en-US'
                if lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia'):
                    langcode = lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia')
                if lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi'):
                    langcode = lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bidi')
                if lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'):
                    langcode = lang_element[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                processed = lang_switch(langcode, processed)
            if is_element:
                if r.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b'):
                    processed = "\\textbf{"+processed+"}"
                if r.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i'):
                    processed = "\\textit{"+processed+"}"
                if r.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u'):
                    processed = "\\underline{"+processed+"}"
            else:
                if r.bold:
                    processed = "\\textbf{"+processed+"}"
                if r.italic:
                    processed = "\\textit{"+processed+"}"
                if r.underline and not in_caption:
                    processed = "\\underline{"+processed+"}"

        else:
            # print('run im citavi-mode')
            if is_element:
                instrText = r.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}instrText')
            else:
                instrText = r.element.xpath('w:instrText')
            try:
                citavi_xml = ET.fromstring(base64.b64decode(instrText[0].text.split()[3]))
                bibtexkey = citavi_xml.findall('.//Entries/Entry/Reference/BibTeXKey')[0].text
                if bibtexkey:
                    processed = "\\cite{"+bibtexkey+"}"
            except:
                print('', end="")

            #     base64 = r.element.xpath('./w:instrText').text.split()[3]
            #     xml = base64.b64decode(base64)
            #     print (xml)# parse xml
            #     processed += "\\cite{{{}}}".format(bibtexkey)

        processed += find_image(r)
        f1, f2 = find_footnote(r, in_caption)
        out += processed + f1
        out2 += f2
    if in_caption:
        return out, out2
    return out

# used by process runs: check if image referenced and handle
def find_image(r):
    try:
        rId = r.element.xpath('w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic/pic:blipFill/a:blip')[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        width = int(r.element.xpath('w:drawing/wp:inline/wp:extent')[0].get('cx')) / 914400
        if width < 4:
            image_insert = image_insert_column
        else:
            image_insert = image_insert_fullwidth
        if rId != 'None':
            filename = rId + r.part.related_parts[rId].filename
            with open (filename, "wb") as imagefile:
                imagefile.write(r.part.related_parts[rId].blob)
        return image_insert.replace('FILENAME', filename)
    except Exception as ex:
        return ''

# used by process runs: check if footnote referenced and handle
def find_footnote(r, in_caption = False):
    try:
        fId = r.element.xpath('w:footnoteReference')[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
        if in_caption:
            return '\\protect\\footnotemark', '\\footnotetext{'+footnotes[fId]+'}'
        return '\\footnote{'+footnotes[fId]+'}', ''
    except Exception as ex:
        return '', ''

# BEGIN work
document = zipfile.ZipFile(sourcefile)
footnotes = {}
tables = {}
# === read footnotes from XML
try:
    xml_content = document.read('word/footnotes.xml')
    doc_xml = ET.fromstring(xml_content)
    for e in doc_xml.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}footnote'):
        if e.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type') == None:
            footnote_id = e.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            footnote_text = process_runs(e.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'))
            footnotes[footnote_id] = footnote_text
except:
    print ('no footnotes')

# The part below is to read tables from the document.xml file.
# It identifies the table, reads out the text inside each cell, produces a table in tex with horizontal lines only.
# the reason it is turned off is that, for lack of time, it was not possible to scale the table to fit the format in tex.
# The table it produces is therefore partly on the page and goes beyond it.
# Further development would require making python identify the dimensions of the table in the xml, make python translate it into a adequate dimensions
# for tex. In other words, there would need to be a module built in this part to "teach" python manually to produce tables to scale for the tex output.
# This could be done, for example, by building two tex table templates (one for in-column tables and one for accross the page tables) which would be capable
# of automatically adjusting their column width and font size depending on whether the table dimensions in xml are above or below a crtain value.
# The second sticky point is table captions who appear above instead of below the table (table captions can be configured in tex)
# and preceded by ":" (ex: ':Tabelle 1 Wertentwicklung des Experiments' instead of 'Tabelle 1: Wertentwicklung des Experiments').
#finally, footnotes in captions must be identified and correctly inserted.
#Good luck!

# === read tables from XML
# try:
#     xml_document = document.read('word/document.xml')                                          #loading document.xml
#     doc_xml2 = ET.fromstring(xml_document                                                      #creating string
#     paragraphs_found = 0                                                                       #creating paragraph class
#     columns_found = 0                                                                          #creating columns class
#     lines_found = 0                                                                            #creating lines class
#     for node in doc_xml2.findall('./*/*'):                                                     #finding all elements in xml
#         if node.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':      #identifying all paragraphs in xml
#             paragraphs_found += 1                                                              #save paragraphs
#         if node.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl':    #identifying all tables in paragraphs
                                        ## TODO: detect and scale dimensions of table. build the two table blueprints
#             columns_found = len(node.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gridCol'))  #counting table columns
#             lines_found = len(node.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))         #counting all table lines
#             table_def = (columns_found*'l')                                                                                #producing zable columns variable for tex
#             lines_def = (lines_found*'toprule')                                                                            #producing table line variable for tex
#             table_text = "\\begin{{table*}}[ht]\\caption{{CAPTION_HERE}}\\centering\\begin{{tabular}}{{{}}}\\toprule \n".format(table_def)    #tex order to produce the table
#             for row in node.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):                                   #finding all table rows
#                 cells = row.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')                                    #finding all table cells
#                 for index, cell in enumerate(cells):                                                                                          #full inventory of table cells
#                     for p in cell.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):                              #finding all paragraphs inside cells
#                         table_text += process_runs(p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'))           #finding all the text inside cells paragraphs
#                     if (index + 1 != len(cells)):                                                                                             #taking out the text in cells
#                         table_text += '& '                                                                                                    #insert cell text into tex preceded by '&'
#                 table_text += '\\\\\\hline \n'                                                                                                #end of row order in tex
#             table_text += '\\end{tabular}\\end{table*}CAPTION_FOOTNOTES_HERE'                                                                 #end of table order in tex and
#             tables[paragraphs_found] = table_text                                                                                             #save table in dictionnary
# except:                                                                                                                                       #what happens if no table is found
#     print ('no table, no table captions found')
# document.close()                                                                                                                              #end of table moducle

# process docx
document = Document(sourcefile)
saveme = ""
for l in open(prefile):
    saveme += l
# replacements in pre tex file

saveme = saveme.replace('BIBFILE', args.bibfile)

# comments in document
parser = argparse.ArgumentParser()
parser.add_argument('--articletype', '-t' , help="not working paper", type= str, default= "Working paper")
parser.add_argument('--history', help="article history", type= str, default= "HISTORY NOT GIVEN")
parser.add_argument('--pub', help="Publication Name", type= str, default= "PUBLICATION NAME AND DATE NOT GIVEN")
parser.add_argument('--institute', help="article history", type= str, default= "INSTITUTE NOT GIVEN")
parser.add_argument('--language', help="article history", type= str, default= "LANGAUGE NOT GIVEN")
args=parser.parse_args([i for i in document.core_properties.comments.split(";") if i])
# more replacements in pre tex file

saveme = saveme.replace('title_STR', document.core_properties.title)
saveme = saveme.replace('pub_STR', args.pub)
saveme = saveme.replace('your_institut_STR', args.institute)
saveme = saveme.replace('type_STR', args.articletype)
saveme = saveme.replace('article_history_STR', args.history)
saveme = saveme.replace('article_language_STR', args.language)



keyword_list = document.core_properties.keywords.split(",")
saveme = saveme.replace('keywords_STR', "\\\\".join(keyword_list))
author_str = ""
author_thx_str = ""
author_list = document.core_properties.author.split(";")
for index, authorelement in enumerate(author_list):
    if authorelement[0] == str('#'):
        author_thx_str = "\\thanks{"+authorelement[1:]+"}"
        author_list.pop(index)
author_str = ", ".join(author_list)
saveme = saveme.replace('AUTHOR_STR', author_str + author_thx_str)

# prepared regexes
heading = re.compile(r"(Heading|Überschrift) (\d+)")
caption = re.compile(r"(Caption|Beschriftung)")
abstract = re.compile(r"Abstract")

outfile = os.path.splitext(os.path.basename(sourcefile))[0] + '.tex'

with open(outfile, "wb") as f:
    current_paragraph = 1
    for p in (document.paragraphs):
        is_heading = heading.match(p.style.name)
        is_caption = caption.match(p.style.name)
        is_abstract = abstract.match(p.style.name)
        if is_heading:
            saveme += '\\' + 'sub'*(int(is_heading.group(2)) - 1) + "section{"+process_runs(p.runs)+"}"
        elif is_caption:
            caption_text, caption_footnotes = process_runs(p.runs, True)
            saveme = saveme.replace('CAPTION_HERE', caption_text)
            saveme = saveme.replace('CAPTION_FOOTNOTES_HERE', caption_footnotes)
        elif is_abstract:
            abstract_text = process_runs(p.runs)
            saveme = saveme.replace('ABSTRACT_HERE', abstract_text)
        else:
            saveme += process_runs(p.runs)
        saveme += "\n"
        # check if we need to insert a table after this paragraph
        if current_paragraph in tables:
            saveme += tables[current_paragraph]
            saveme += "\n"
        current_paragraph += 1

    # finalize: if we have placeholders still left, remove
    saveme = saveme.replace('CAPTION_HERE', "")
    saveme = saveme.replace('CAPTION_FOOTNOTES_HERE', "")
    saveme = saveme.replace('ABSTRACT_HERE', "")
    # write output to file
    f.write((saveme+'\n').encode('UTF-8'))
    f.writelines([l.encode('UTF-8') for l in open(postfile)])

print ('All done, output file {} written.'.format(outfile))
# END
