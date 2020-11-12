import sys

# first install libreoffice, like:
# sudo apt-get install libreoffice --no-install-recommends

# second start libreoffice --socket, like:
# soffice_bin=/usr/lib/libreoffice/program/soffice.bin
# $soffice_bin --headless --nologo --nofirststartwizard --norestore --nodefault --accept='socket,host=localhost,port=8100;urp;StarOffice.Service'

# import uno system package; so far (11.12.2020)'uno' is not in 'pip'
# to check if it's installed:
# $ python3
# >>> import uno
# >>> dir(uno)
# >>> print(uno.__file__)

uno_package_location = '/usr/lib/python3/dist-packages/'
sys.path.append(uno_package_location)
import uno

from os.path import isfile, join
from os import getcwd
from unohelper import systemPathToFileUrl, absolutize
from com.sun.star.beans import PropertyValue


outputfile = "helloworld.odt"
url = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
ctxLocal = uno.getComponentContext()
smgrLocal = ctxLocal.ServiceManager
resolver = smgrLocal.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", ctxLocal )
ctx = resolver.resolve( url )
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx )

cwd = systemPathToFileUrl( getcwd() )
destFile = absolutize( cwd, systemPathToFileUrl(outputfile) )
inProps = PropertyValue( "Hidden" , 0 , True, 0 ),
newdoc = desktop.loadComponentFromURL(
    "private:factory/swriter", "_blank", 0, inProps )
#get the XText interface
text   = newdoc.Text
#create an XTextRange at the end of the document
tRange = text.End
#and set the string
tRange.String = "Hello World (in Python)"
print("Hello World (in Python)")
newdoc.storeAsURL(destFile, ())
newdoc.dispose()
print(f"File {outputfile} created.")

# run this file with `python3 helloworld.py`
# then close libreoffice-socket: kill PID