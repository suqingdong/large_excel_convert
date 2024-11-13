from lxml.etree import XMLPullParser


class XMLParser(object):
    def __init__(self, xml_file):
        self.xml_file = xml_file
        self.parser = XMLPullParser(['start', 'end'])