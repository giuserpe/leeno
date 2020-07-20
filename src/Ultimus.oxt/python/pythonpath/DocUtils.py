'''
try to overcome limits on document user defined attributes
'''
from com.sun.star.beans import PropertyAttribute
from com.sun.star.beans import PropertyExistException, UnknownPropertyException
from com.sun.star.lang import IllegalArgumentException
import PersistUtils

def setDocUserDefinedAttribute(oDoc, name, value):
    userProps = oDoc.DocumentProperties.UserDefinedProperties

    # stringize the property
    sValue = PersistUtils.var2string(value)

    # try to add the property, exception if exists
    try:
        userProps.addProperty(name, PropertyAttribute.REMOVABLE, sValue)
    except PropertyExistException:
        pass
    try:
        userProps.setPropertyValue(name, sValue)
    except IllegalArgumentException:
        try:
            userProps.removeProperty(name)
            userProps.addProperty(name, PropertyAttribute.REMOVABLE, sValue)
            userProps.setPropertyValue(name, sValue)
        except Exception:
            pass

def getDocUserDefinedAttribute(oDoc, name):
    userProps = oDoc.DocumentProperties.UserDefinedProperties

    try:
        sValue = userProps.getPropertyValue(name)
        return PersistUtils.string2val(sValue)

    except Exception:
        return None

def hasDocUserDefinedAttribute(oDoc, name):
    return getDocUserDefinedAttribute(oDoc, name) != None

def removeDocUserDefinedAttribute(oDoc, name):
    userProps = oDoc.DocumentProperties.UserDefinedProperties

    # try to add the property, exception if exists
    try:
        userProps.removeProperty(name)
    except UnknownPropertyException:
        pass


def storeDataBlock(oDoc, baseName, data):
    '''
    baseName : nome base per il blocco dati. Viene preposto ai dati prima di salvarli
    data : un dizionario contenente una serie di chiave:valore. Le chiavi devono essere stringhe
    '''
    for key, value in data.items():
        setDocUserDefinedAttribute(oDoc, baseName + '.' + key, value)


def loadDataBlock(oDoc, baseName):
    '''
    baseName : prefisso per il blocco di dati richiesto
    Vengono lette TUTTE le proprietà inizianti con baseName
    e restituite sotto forma di dizionario { key: value... }
    '''
    userProps = oDoc.DocumentProperties.UserDefinedProperties
    props = userProps.PropertySetInfo.Properties
    res = {}
    for prop in props:
        if prop.Name.startswith(baseName + '.'):
            name = prop.Name[len(baseName) + 1:]
            val = PersistUtils.string2var(userProps.getPropertyValue(prop.Name))
            if val is not None:
                res[name] = val
    return res
