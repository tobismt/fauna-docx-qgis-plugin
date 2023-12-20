import pandas as pd
import docx
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, Cm
import os

class redListFauna:
    def __init__(self, fauna_layer, outpath):
        self.fauna_layer = fauna_layer
        self.outpath = outpath
        self.basepath = os.path.dirname(os.path.realpath(__file__))
        self.LUT = pd.read_csv(f"{self.basepath}/fauna.csv", sep='|')
        self.legend = pd.read_csv(f"{self.basepath}/legend.csv", sep='|')
        self.doc = docx.Document()
        
        self.get_arten_list(self.fauna_layer)
        self.create_df()
        self.add_header()
        self.df_to_word()
        self.color_cells(self.doc.tables[0])
        self.center_text()
        self.create_legend()
        self.save()
        
    def add_header(self):
        header = self.doc.sections[0].header
        header.paragraphs[0].text = "Rote-Liste Fauna im Untersuchungsgebiet\n"
        header.paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        style = self.doc.styles['Heading 1']
        font = style.font
        font.size = Pt(16)
        
        header.paragraphs[0].style = self.doc.styles['Heading 1']
            
    def get_arten_list(self, lyr):
        cols = [f.name() for f in lyr.fields()]
        datagen = ([f[col] for col in cols] for f in lyr.getFeatures())

        df = pd.DataFrame.from_records(data=datagen, columns=cols)
        
        list = df.Name.unique()
        self.list = list


    def create_df(self):
        df = pd.DataFrame(self.list, columns = ['Name'])
        merge = df.merge(self.LUT, how='left', left_on='Name', right_on='Name')
        merge = merge[['Name', 'Deutscher Name', 'aktuelle Bestandssituation', 'kurzfristiger Bestandstrend', 'langfristiger Bestandstrend', 'RL Kat.']]
        merge.columns = [i.title() for i in merge.columns]
        merge = merge.fillna('-')
        merge = merge.sort_values('Name')
        
        self.df = merge

    def df_to_word(self):
        t = self.doc.add_table(self.df.shape[0]+1, self.df.shape[1])
        for j in range(self.df.shape[-1]):
            t.cell(0,j).text = self.df.columns[j]

        for i in range(self.df.shape[0]):
            for j in range(self.df.shape[-1]):
                t.cell(i+1,j).text = str(self.df.values[i,j])

        for i, col in enumerate(t.columns):
            if i == 0:
                col.width = Cm(4)
            else:
                col.width = Cm(2.5)

    def color_cells(self, table):
        num = 0
        elems = []
        for row in table.rows:
            for cell in row.cells:
                if '0' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#3ec902"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if '1' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#80c902"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if '2' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#c9b202"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if '3' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#a30202"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if 'G' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#fa7000"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if 'R' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#b300fa"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if 'V' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#2302c9"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if 'D' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#acacad"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if '*' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="#acadad"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1
                if '♦' == cell.text:
                    elems.append(parse_xml(r'<w:shd {} w:fill="9edd23"/>'.format(nsdecls('w'))))
                    cell._tc.get_or_add_tcPr().append(elems[num])
                    num += 1

    def create_legend(self):
        self.doc.add_paragraph('')
        self.legend.fillna('', inplace=True)
        t = self.doc.add_table(self.legend.shape[0] + 1, self.legend.shape[1])
        for j in range(self.legend.shape[-1]):
            t.cell(0, j).text = self.legend.columns[j]

        for i in range(self.legend.shape[0]):
            for j in range(self.legend.shape[-1]):
                if '-' not in str(self.legend.values[i, j]):
                    t.cell(i + 1, j).text = str(self.legend.values[i, j])

        t.alignment = WD_TABLE_ALIGNMENT.CENTER

        t.autofit = False
        t.allow_autofit = False
        for col in t.columns:
            col.width = Cm(1.75)

        for i, row in enumerate(t.rows):
            if i == 0:
                row.cells[0].merge(row.cells[1])
                row.cells[0].text = 'Rote Liste Status'
                row.cells[2].merge(row.cells[3])
                row.cells[3].text = 'Aktuelle Bestandssituation'
                row.cells[4].merge(row.cells[5])
                row.cells[5].text = 'Bestandstrend langfristig'
                row.cells[6].merge(row.cells[7])
                row.cells[7].text = 'Bestandstrend kurzfristig'
            for cell in row.cells:
                if '-' in cell.text:
                    cell.text = ''
                paragraphs = cell.paragraphs
                for paragraph in paragraphs:
                    for run in paragraph.runs:
                        font = run.font
                        if row.cells[0].text == 'Rote Liste Status':
                            font.bold = True
                        font.size = Pt(6)
                    paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER



                    
    def center_text(self):
        for row in self.doc.tables[0].rows:
            if row.cells[1].text == 'Deutscher Name':
                for cell in row.cells:
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        for run in paragraph.runs:
                            font = run.font
                            font.size = Pt(12)
                            font.bold = True
            for cell in row.cells:
                    paragraphs = cell.paragraphs
                    for paragraph in paragraphs:
                        paragraph.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
            
        
    def save(self):
        self.doc.save(self.outpath)
        