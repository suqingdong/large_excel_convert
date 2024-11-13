import zipfile

from lxml import etree


class ExcelParser(object):
    def __init__(self, file, tag='row', sheet=1, chunksize=1024*1024):
        self.sheet = sheet
        self.chunksize = chunksize
        self.zip_handle = zipfile.ZipFile(file)
        self.namespace = self.get_namespace()
        self.tag = f'{{{self.namespace}}}{tag}' if self.namespace else tag

    def get_namespace(self):
        tree = etree.parse(self.zip_handle.open('xl/workbook.xml'))
        return list(tree.getroot().nsmap.values())[0]

    def rows(self):
        xml_parser = etree.XMLPullParser(events=('end',), tag=self.tag)

        with self.zip_handle.open(f'xl/worksheets/sheet{self.sheet}.xml') as xml:
            for chunk in iter(lambda: xml.read(self.chunksize), b''):
                xml_parser.feed(chunk)
                yield from self.read_events(xml_parser)
        yield from self.read_events(xml_parser)

    def read_events(self, xml_parser):
        for action, elem in xml_parser.read_events():
            yield list(self.get_row_values(elem))
            elem.clear()


    @staticmethod
    def get_row_values(elem):
        for child in elem.getchildren():
            text = '|'.join(t.strip().replace('\n', ';') for n, t in enumerate(child.itertext()) if t.strip())
            yield text
