from com.sun.star.beans import PropertyAttribute
from com.sun.star.beans import PropertyExistException, UnknownPropertyException

def setDocUserDefinedAttribute(oDoc, name, value):
    userProps = oDoc.DocumentProperties.UserDefinedProperties

    # try to add the property, exception if exists
    try:
        userProps.addProperty(name, PropertyAttribute.REMOVABLE, '')
    except PropertyExistException:
        pass
    userProps.setPropertyValue(name, value)

def getDocUserDefinedAttribute(oDoc, name):
    userProps = oDoc.DocumentProperties.UserDefinedProperties

    try:
        return userProps.getPropertyValue(name)
    except UnknownPropertyException:
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
