'''
some not-too-complicated persistence methods
'''
from datetime import date
import LeenoUtils

_TYPEMAP = {
    str: lambda x : '(str)' + x,
    int: lambda x : '(int)' + str(x),
    float: lambda x : '(float)' + str(x),
    date: lambda x : '(date)' + LeenoUtils.date2String(x, 1),
    bool: lambda x : '(bool)' + str(x)
}

_TYPENAMEMAP = {
    'str': str,
    'int': int,
    'float': float,
    'date': LeenoUtils.string2Date,
    'bool': lambda x : True if x == 'True' else False,
}


def string2var(s):
    '''
    convert the string-format data (type)xxxxxx to
    its value equivalent
    examples :  (int)5 --> 5
                (date)25/02/2020 --> date(2020, 02, 25)
    '''
    if not s.startswith('('):
        return None
    closePos = s.find(')')
    if closePos < 0:
        return None
    typ = s[1:closePos]
    if not typ in _TYPENAMEMAP:
        return None
    return _TYPENAMEMAP[typ](s[closePos+1:])

def var2string(var):
    '''
    convert the variable var to a storable string
    '''
    if not type(var) in _TYPEMAP :
        return None
    return _TYPEMAP[type(var)](var)
