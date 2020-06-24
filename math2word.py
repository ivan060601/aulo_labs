import latex2mathml.converter
from lxml import etree

def math_to_word(input_in_latex):
    # Creates mathml string
    mathml_string = latex2mathml.converter.convert(input_in_latex)
    # Converts mathml string
    tree = etree.fromstring(mathml_string)
    xslt = etree.parse("D:/PyCharm Projects/Project1/res/MML2OMML.XSL")
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()