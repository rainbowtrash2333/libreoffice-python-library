import sys
import time
import subprocess
from .config import config
# uno.py directory (package) path
sys.path.append(config['libreoffice']['python_uno_location'])
# ***** UNO *****
import uno
NoConnectException = uno.getClass('com.sun.star.connection.NoConnectException')
# from com.sun.star.connection import NoConnectException
import socket

now = time.time


class LOprocess:
    def __init__(self,
                 connection=config['libreoffice']['connection']):

        # command flags array
        self.flags = config['libreoffice']['flags']

        # libreoffice binary location
        self.libreoffice_bin = connection['binary_location']
        # # UNO socket connection settings
        self.host = connection['host']
        self.port = connection['port']
        # will become content of '--accept=' flag, like f"--accept='{accept_open}'"
        # something like "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"
        self.accept_open = connection['accept_open']
        # this is used when connetion to running libreoffice process
        # something like "uno:socket,host=%s,port=%s,tcpNoDalay=1;urp;StarOffice.ComponentContext"
        self.connection_url = connection['connection_url']
        self.timeout = connection['timeout']

        # process ref of running libreoffice set from startup method
        self.loproc = None
        self.desktop = None
        self.__is_start = False


    def startup(self):
        """ Starts libreoffice process.
        """
        accept_open = self.accept_open % (self.host, self.port)
        # "--accept=%s" => %s is replaced with accept_open

        start_command_flags = [s.replace("%s", accept_open) for s in self.flags]

        start_command =  [self.libreoffice_bin, *start_command_flags]

        # shell=False Cross-platform security
        lo_proc = subprocess.Popen(start_command,shell=False)
        self.__is_start = True
        return lo_proc

    def connect(self):
        if not self.__is_start:
            self.startup()
        local_ctx = uno.getComponentContext()
        smgr_local = local_ctx.ServiceManager
        resolver = smgr_local.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx)
        url = self.connection_url % (self.host,  self.port)

        uno_ctx = None
        # try to resolve connection
        try:
            # url = "uno:socket,host=localhost,port=2002,tcpNoDalay=1;urp;StarOffice.ComponentContext"
            # url = "uno:pipe,name=somepipename;urp;StarOffice.ComponentContext"
            uno_ctx = resolver.resolve(url)
        except NoConnectException:
            print(f"connection failed: Make sure soffice is running and port {self.port} is available")
            exit(1)


        return uno_ctx

    def terminate(self, desktop):
        desktop.terminate()

