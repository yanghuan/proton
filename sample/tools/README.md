[English](https://github.com/sy-yanghuan/proton#proton)   [Chinese](https://github.com/sy-yanghuan/proton#proton-1)  
# proton
Proton is a excel export configuration file for the tool, you can export to xml, json, lua format, through external expansion can automatically generate the configuration to read the code, simple and flexible easy to use, indeed powerful.

## Features
- Writeing in Python,cross-platform, referenced [xlrd](http://www.lexicon.net/sjmachin/xlrd.html) only, [full code](https://github.com/sy-yanghuan/proton/blob/master/proton.py), just more than 500 lines  
- Has a specific rule syntax description excel format information, simple and easy to understand, flexible and powerful, [detailed description](https://github.com/sy-yanghuan/proton/wiki/document_en) 
- Can export excel format information for external use, can be used to automatically generate read configuration code

## Generates an auto-read code
Use the "-c" parameter to generate a json file containing excel format information, each language can be automatically generated to achieve this code to read the tool, [the specific format](https://github.com/sy-yanghuan/proton/wiki/schema_en). Has achieved the C # language tools, other language users, can be realized, welcomed the realization of the code links for the needs of people to use.
- [CSharpGeneratorForProton](https://github.com/sy-yanghuan/CSharpGeneratorForProton) generates C # code that reads xml, json, protobuf. You can convert xml, json to protobuf's binary format and generate the corresponding read code (using protobuf-net).

## Command Line Parameters
```cmd
usage python proton.py [-p filelist] [-f outfolder] [-e format]
Arguments 
-p      : input excel files, use space to separate 
-f      : out folder
-e      : format, json or xml or lua     

Options
-s      ：sign, controls whether the column is exported, defalut all export
-t      : suffix, export file suffix
-c      : a file path, save the excel structure to json, 
          the external program uses this file to automatically generate the read code      
-h      : print this help message and exit
```

## Documentation
Wiki https://github.com/sy-yanghuan/proton/wiki/document_en  
FAQ https://github.com/sy-yanghuan/proton/wiki/FAQ_en

## *License*
[Apache 2.0 license](https://github.com/sy-yanghuan/proton/blob/master/LICENSE).

_____________________
# proton
proton是一个将excel导出为配置文件的工具，可以导出为xml、json、lua格式，通过外部扩展可支持自动生成读取配置的代码，简单灵活易于使用，确不失强大。
## 特点
- python编写可跨平台使用，仅依赖第三方库[xlrd](http://www.lexicon.net/sjmachin/xlrd.html)，[完整代码仅500余行](https://github.com/sy-yanghuan/proton/blob/master/proton.py)。
- 有特定的规则语法描述excel的格式信息，简洁易懂，灵活强大，[详细说明](https://github.com/sy-yanghuan/proton/wiki/document_zh)。
- 可导出excel格式信息供外部程序使用，可用来自动生成读取配置的代码。

##后端程序（生成自动读取的代码）
使用“-c”参数可生成内含excel格式信息的json文件，各个语言可据此实现自动生成读取代码的工具，[具体格式说明](https://github.com/sy-yanghuan/proton/wiki/schema_zh)。已经实现了C#语言的工具，其他语言使用者，可自行实现，欢迎提供实现的代码链接，以供需要的同学使用。

- [CSharpGeneratorForProton](https://github.com/sy-yanghuan/CSharpGeneratorForProton) 可生成读取xml、json、protobuf的C#代码。 可将xml、json转换为protobuf的二进制格式，并生成对应的读取代码（使用protobuf-net）。

## 命令行参数
```cmd
usage python proton.py [-p filelist] [-f outfolder] [-e format]
Arguments 
-p      : input excel files, use space to separate 
-f      : out folder
-e      : format, json or xml or lua     

Options
-s      ：sign, controls whether the column is exported, defalut all export
-t      : suffix, export file suffix
-c      : a file path, save the excel structure to json, 
          the external program uses this file to automatically generate the read code      
-h      : print this help message and exit
```
##文档
格式说明 https://github.com/sy-yanghuan/proton/wiki/document_zh  
FAQ https://github.com/sy-yanghuan/proton/wiki/FAQ_zh

##*许可证*
[Apache 2.0 license](https://github.com/sy-yanghuan/proton/blob/master/LICENSE).


