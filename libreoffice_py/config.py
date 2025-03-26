import toml


def _get_lo_config():
    lo_config_default = r"""

      [libreoffice]
      # To find this location, open python3 interactive, >>> import uno; print(uno.__file__)
      python_uno_location = "/usr/lib/python3/dist-packages/"
      # windows location
      # python_uno_location = "C:\\Program Files\\LibreOffice\\program\\"
      
      # soffice.bin should be run with these flags
      flags = ["--headless", "--invisible", "--nocrashreport", "--nodefault", "--nofirststartwizard", "--nologo", "--norestore", "--accept=%s"]

      [libreoffice.connection]
      binary_location = "/usr/lib/libreoffice/program/soffice.bin"
      # windows binary_location
      # binary_location = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
      
      # UNO socket connection settings
      host = "localhost"
      port = 2002

      # content of '--accept=' flag for starting soffice.bin with open connection
      accept_open = "socket,host=%s,port=%s,tcpNoDelay=1;urp;StarOffice.ComponentContext"

      # this is used to connect running libreoffice process with opened socket
      connection_url = "uno:socket,host=%s,port=%s,tcpNoDalay=1;urp;StarOffice.ComponentContext"

      # how much time to wait for connecting
      timeout = 5
      """
    config_path = './lo_config.toml'

    # load connection settings from Config toml file
    try:
         return toml.load(config_path)
    except FileNotFoundError:
        # creating default libreoffice config file
        print(f"Creating default config file: {config_path}")
        with open(config_path, 'w') as f:
            f.write(lo_config_default)
        return toml.loads(lo_config_default)
    except toml.TomlDecodeError:
        print("config file encoding error，use default config file")
        return toml.loads(lo_config_default)


# get config
config = _get_lo_config()

