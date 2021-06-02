from lxml import etree as ET
import re
from pprint import pprint

dummy_path = r'sample.twb'
regex_xml_formatter = re.compile('\\n\s+\w')


class TableauObjectParser:

    @staticmethod
    def _parse_name(name):
        """
        Converts a xml attribute to be used as a reference to snake case
        :param name:
        :return:
        """
        cleansed_name = name.translate(str.maketrans({'[': '', ']': '', '-': '', })).lower()
        spaced_name = ' '.join(cleansed_name.split())
        return spaced_name.replace(' ', '_')

    @staticmethod
    def parse_datasource(container):
        try:
            return TableauObjectParser._parse_name(container.xml.attrib['caption'])
        except KeyError:
            return TableauObjectParser._parse_name(container.xml.attrib['name'])

    @staticmethod
    def parse_parameter(container):
        return TableauObjectParser._parse_name(container.xml.attrib['name'])

    @staticmethod
    def parse_calculated_field(container):
        return TableauObjectParser._parse_name(container.xml.attrib['caption'])

    @staticmethod
    def parse_default(container):
        return TableauObjectParser._parse_name(container.xml.tag)


class XMLContainer:

    def __init__(self, xml, workbook, parent=None):
        self._workbook = workbook
        self.xml = xml
        self.parent = parent
        self.expose_attributes()
        self.reference_parser_func = self._identify_parser_func()
        self.children = []

        if parent:
            self.attach_to_parent()
            self.abs_xpath = f'{self.parent.abs_xpath}/{self.xml.tag}'
        else:
            self.abs_xpath = f'/workbook/{self.xml.tag}'

        if len(self.xml) > 0:
            self.children = self.parse_children()
        if len(self.xml) == 1:
            self.child_attributes = []

    def attach_to_parent(self):
        """
        Attaches the child instance to its parent via the parent's setattr method
        :return:
        """

        if self.parent is None:
            print(self.xml.tag, ' has no parent')
        else:
            self.parent.__setattr__(self.reference_parser_func(self), self)

    def expose_attributes(self, _dict=None, override_existing_keys=True):
        """
        exposes an object's attributes
        :return:
        """

        if _dict is None:
            for k, v in self.xml.attrib.items():
                self.__setattr__(k, v)
        else:
            for k, v in _dict.items():
                if override_existing_keys:
                    self.__setattr__(k, v)
                else:
                    try:
                        self.__dict__[k]
                    except KeyError:
                        self.__setattr__(k, v)

    def _identify_parser_func(self):

        if self.parent is None:
            pass
        elif self.parent.xml.tag == 'datasource':
            if self.xml.tag == 'column':
                if self.parent.xml.attrib['name'] == 'Parameters':
                    return TableauObjectParser.parse_parameter
                if 'Calculation' in self.xml.attrib['name'] and self.parent.xml.attrib['name'] != 'Parameters':
                    return TableauObjectParser.parse_calculated_field

        elif self.xml.tag == 'datasource':
            return TableauObjectParser.parse_datasource

        return TableauObjectParser.parse_default

    def parse_children(self):
        """
        Generates child containers and references for the parent as necessary based on the xml specification
        :return:
        """

        children = []

        for child in self.xml:
            child_container = XMLContainer(child, self._workbook, self)
            children.append(child_container)

            if len(self.xml) == 1:
                self.expose_attributes(_dict=child_container.xml.attrib, override_existing_keys=False)
                self.child_attributes = [x for x in child_container.xml.attrib.keys()]

        return children

    def __repr__(self):
        return str(ET.dump(self.xml))

    def update(self):
        for x in self._workbook.xml.xpath(self.abs_xpath):
            print(x.tag)


class Workbook(XMLContainer):

    def __init__(self, workbook_path, shortcuts=True, shortcut_table=None):
        super().__init__(ET.parse(workbook_path).getroot(), self)
        self.shortcut_table = shortcut_table

        if shortcuts:
            self.apply_shortcuts()

    def apply_shortcuts(self):
        if self.shortcut_table is None:
            self.shortcut_table = dict()

            for _k, _v in self.datasources.__dict__.items():
                if isinstance(_v, XMLContainer):
                    self.shortcut_table[_k] = _v

        for k, v in self.shortcut_table.items():
            self.__setattr__(k, v)

    def save_workbook(self, path):
        _root = ET.Element(self.xml.tag, self.xml.attrib, self.xml.nsmap)

        container_stack = [self]
        element_stack = [_root]
        while container_stack:
            container = container_stack.pop()
            parent_element = element_stack.pop()
            if len(container.xml) > 0:
                for child_container in container.children:
                    child_element = ET.Element(child_container.xml.tag, child_container.xml.attrib)
                    if child_container.xml.text:
                        if re.match(regex_xml_formatter, child_container.xml.text) is None and '\n' in child_container.xml.text:
                            child_container.xml.text = None
                        else:
                            child_element.text = child_container.xml.text
                    parent_element.append(child_element)
                    element_stack.append(child_element)
                    container_stack.append(child_container)

        with open(path, 'wb') as doc:
            doc.write(ET.tostring(_root, pretty_print=True, xml_declaration=True, encoding='utf-8'))
        return _root

    def __repr__(self):
        return str(ET.dump(self.xml))

    def get_calculations(self):
        container_stack = [self]
        data_source_calculations = dict()
        parameter_calculations = dict()

        while True:
            container = container_stack.pop()
            if len(container.children) > 0:
                container_stack.extend(container.children)
            else:
                try:
                    if container.parent.parent.name == 'Parameters':
                        parameter_calculations[container.parent.caption] = container.formula
                    else:
                        data_source_calculations[container.parent.caption] = container.formula
                except AttributeError:
                    pass

            if len(container_stack) == 0:
                return parameter_calculations, data_source_calculations


"""
sample = Workbook(workbook_path=dummy_path)
pprint(sample.get_calculations())
print(sample.sample_eu_superstore.profit_ratio.formula)
print(sample.sample_eu_superstore.profit_ratio.calculation)
"""

# TODO update a calc using update()
#sample.sample_eu_superstore.profit_ratio.formula.update('5 * 5')
# TODO bulk update based on a config file i.e TOML (accepts a dictionary?)
