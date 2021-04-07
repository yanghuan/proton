#encoding=utf-8

# Need to export the public configuration file (client, server needs) 需要导出的公共配置文件(客户端,服务器都需要)
EXPORT_FILES = [
"hero.xlsx",
"mount.xlsx",
]

# Additional configuration files that the client needs to export (only the client needs) 客户端额外需要导出的额外配置文件(仅客户端需要)
EXPORT_CLIENT_ONLY = [
"text.xlsx"
]

# Server-side need to export the configuration file (only the server needs) 服务器端额外需要导出的配置文件(仅服务器需要)
EXPORT_SERVER_ONLY = [
]

# do not modify the following

import os
import platform
import traceback
import shutil
import sys

exportscript = '../proton.py'     
pythonpath = 'tools\\py37\\py37.exe ' if platform.system() == 'Windows' else 'python '

class ExportError(Exception):
  pass

def export(filelist, format, sign, outfolder, suffix, schema):
  cmd = r' -p "' + ','.join(filelist) + '" -f ' + outfolder + ' -e ' + format + ' -s ' + sign
  if suffix:
    cmd += ' -t ' + suffix
  if schema:
    cmd += ' -c ' + schema
  cmd = pythonpath + exportscript + cmd
  code = os.system(cmd)
  if code != 0:
    raise ExportError('export excel fail, please see print')

def codegenerator(schema, outfolder, namespace, suffix):
  if os.path.exists(schema):
    cmd = 'tools\CSharpGeneratorForProton\CSharpGeneratorForProton.exe ' + '-n ' + namespace + ' -f ' + outfolder + ' -p ' + schema
    if suffix:
      cmd += ' -t ' + suffix 
    code = os.system(cmd)
    os.remove(schema)      
    if code != 0:
      raise ExportError('codegenerator fail, please see print')
        
def exportserver():
  export(EXPORT_FILES + EXPORT_SERVER_ONLY, 'json', 'server', 'config_server', 'Config', 'schemaserver.json')
  codegenerator('schemaserver.json', 'config_server/ConfigGenerator/Template', 'Ice.Project.Config', 'Template') 
    
def exportclient():
  export(EXPORT_FILES + EXPORT_CLIENT_ONLY, 'lua', 'client', 'config_client', 'Template', None)
    
def main():
  try:
    exportserver()
    exportclient()
    print("all operation finish successful")
    return 0
  except ExportError as e:
    print(e)
    print("has error, see logs, please return key to exit")
    input()
    return 1
  except Exception as e:
    traceback.print_exc()
    print("has error, see logs, please return key to exit")
    input()
    return 1
    
if __name__ == '__main__':
    sys.exit(main())
