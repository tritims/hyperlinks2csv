import zipfile
import re
import json
import base64
from docx import Document
from os.path import basename
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from bs4 import BeautifulSoup
import pandas as pd

def __get_linked_text(soup):

    links = []

    # This kind of link has a corresponding URL in the _rel file.
    for tag in soup.find_all("hyperlink"):
        # try/except because some hyperlinks have no id.
        try:
            links.append({"id": tag["r:id"], "text": tag.text})
        except:
            pass

    # This kind does not.
    for tag in soup.find_all("instrText"):
        # They're identified by the word HYPERLINK
        if "HYPERLINK" in tag.text:
            # Get the URL. Probably.
            url = tag.text.split('"')[1]

            # The actual linked text is stored nearby tags.
            # Loop through the siblings starting here.
            temp = tag.parent.next_sibling
            text = ""

            while temp is not None:
                # Text comes in <t> tags.
                maybe_text = temp.find("t")
                if maybe_text is not None:
                    # Ones that have text in them.
                    if maybe_text.text.strip() != "":
                        text += maybe_text.text.strip()

                # Links end with <w:fldChar w:fldCharType="end" />.
                maybe_end = temp.find("fldChar[w:fldCharType]")
                if maybe_end is not None:
                    if maybe_end["w:fldCharType"] == "end":
                        break

                temp = temp.next_sibling

            links.append({"id": None, "href": url, "text": text})

    return links


def __get_links(file_name):
    document = Document(file_name)
    rels = document.part.rels
    links = []
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            links.append({"id": rel, "url": rels[rel]._target})      
    return links

def hyperlinks2csv(file_name, target_csv):
    # file_name="X.com.docx" # Give the docx file name with extension
    archive = zipfile.ZipFile(file_name, "r")
    file_data = archive.read("word/document.xml")
    doc_soup = BeautifulSoup(file_data, "xml")
    linked_text = __get_linked_text(doc_soup)
    linked_links = __get_links(file_name)
    df_text = pd.DataFrame(linked_text)
    df_links = pd.DataFrame(linked_links)
    df = pd.merge(df_text, df_links, on='id')
    df[['text', 'url']].to_csv(target_csv)
    