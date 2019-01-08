from networkx.utils import open_file
from networkx.exception import NetworkXError
import re

try:
    long
except NameError:
    long = int
try:
    unicode
except NameError:
    unicode = str

def generate_gml(G, stringizer=None):
 
    valid_keys = re.compile('^[A-Za-z][0-9A-Za-z]*$')

    def stringize(key, value, ignored_keys, indent, in_list=False):
        if not isinstance(key, (str, unicode)):
            raise NetworkXError('%r is not a string' % (key,))
        if not valid_keys.match(key):
            raise NetworkXError('%r is not a valid key' % (key,))
        if not isinstance(key, str):
            key = str(key)
        if key not in ignored_keys:
            if isinstance(value, (int, long, bool)):
                if key == 'nxlabel':
                    yield indent + key + ' "' + str(value) + '"'
                elif value is True:
                    # python bool is an instance of int
                    yield indent + key + ' 1'
                elif value is False:
                    yield indent + key + ' 0'
                else:
                    yield indent + key + ' ' + str(value)
            elif isinstance(value, float):
                text = repr(value).upper()
                # GML requires that a real literal contain a decimal point, but
                # repr may not output a decimal point when the mantissa is
                # integral and hence needs fixing.
                epos = text.rfind('E')
                if epos != -1 and text.find('.', 0, epos) == -1:
                    text = text[:epos] + '.' + text[epos:]
                if key == 'nxlabel':
                    yield indent + key + ' "' + text + '"'
                else:
                    yield indent + key + ' ' + text
            elif isinstance(value, dict):
                yield indent + key + ' ['
                next_indent = indent + '  '
                for key, value in value.items():
                    for line in stringize(key, value, (), next_indent):
                        yield line
                yield indent + ']'
            elif isinstance(value, (list, tuple)) and key != 'nxlabel' \
                    and value and not in_list:
                next_indent = indent + '  '
                for val in value:
                    for line in stringize(key, val, (), next_indent, True):
                        yield line
            else:
                if stringizer:
                    try:
                        value = stringizer(value)
                    except ValueError:
                        raise NetworkXError(
                            '%r cannot be converted into a string' % (value,))
                if not isinstance(value, (str, unicode)):
                    raise NetworkXError('%r is not a string' % (value,))
                yield indent + key + ' "' + escape(value) + '"'

    multigraph = G.is_multigraph()
    yield 'graph ['

    # Output graph attributes
    if G.is_directed():
        yield '  directed 1'
    if multigraph:
        yield '  multigraph 1'
    ignored_keys = {'directed', 'multigraph', 'node', 'edge'}
    for attr, value in G.graph.items():
        for line in stringize(attr, value, ignored_keys, '  '):
            yield line

    # Output node data
    node_id = dict(zip(G, range(len(G))))
    print(node_id)
    ignored_keys = {'nxid', 'nxlabel'}
    for node, attrs in G.nodes.items():
        yield '  node ['
        yield '    nxid ' + str(node_id[node])
        for line in stringize('nxlabel', node, (), '    '):
            yield line
        for attr, value in attrs.items():
            for line in stringize(attr, value, ignored_keys, '    '):
                yield line
        yield '  ]'

    # Output edge data
    ignored_keys = {'nxsource', 'nxtarget'}
    kwargs = {'data': True}
    if multigraph:
        ignored_keys.add('key')
        kwargs['keys'] = True
    for e in G.edges(**kwargs):
        yield '  edge ['
        yield '    source ' + str(e[0])
        yield '    target ' + str(e[1])
        if multigraph:
            for line in stringize('key', e[2], (), '    '):
                yield line
        for attr, value in e[-1].items():
            for line in stringize(attr, value, ignored_keys, '    '):
                yield line
        yield '  ]'
    yield ']'


@open_file(1, mode='wb')
def write_gml(G, path, stringizer=None):

    for line in generate_gml(G, stringizer):
        path.write((line + '\n').encode('ascii'))
        

def escape(text):
    """Use XML character references to escape characters.

    Use XML character references for unprintable or non-ASCII
    characters, double quotes and ampersands in a string
    """
    def fixup(m):
        ch = m.group(0)
        return '&#' + str(ord(ch)) + ';'

    text = re.sub('[^ -~]|[&"]', fixup, text)
    return text if isinstance(text, str) else str(text)
