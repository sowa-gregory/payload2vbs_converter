# Payload2VBSConverter

Tool to embed payload (eg. malicious executable file) inside Microsoft office applications.

## Description

Presented tool allows for convertion of any type of the file into vbs script. Input file is encoded in base64 algorithm and represented as string constants. Generated output vbs code is embeddable into one of the MS office applications (Word, Excep, PowerPoint). Output script contains functions to decode payload and store it back into file.

This software doesn't contain any magical method to deliver payload to the target machine. It is up to the programmer to prepare office document (for instance Excel sheet) containing something valuable for the potential user. Afterwards, this excel can be used as method to deliver additional payload to remote workstation, as for example email attachment. VBS macros build into Excel file could then invoke generated script to store transported payload into filesystem of targeted victim.

## Disclaimer

This is a not hacking tool and can be used for educational purposes only.

## Usage

The tool is written in Python version 2. Only standard build-in libraries are used, no additional modules must be installed. Run the tool by following command line:

```shell
python payload2vbs_converter.py [filename]
```

### Arguments

- **filename** - path to file to be embedded inside generated output script.

## VBS limitations

Excel VBS limits size of individual string used as function parameters. If large amount of strings are presented inside the script, it Excel prevents from saving worksheet content. Therefore generated script uses relatively short string constants.

Another limitation is related to maximum number of lines in one procedure. The size of individual code blocks must be controlled. The tool automatically splits generated code into many separated procedures.

## Output code structure

Generated output code contains three main sections:

- base64 decoder and file storage functions

Internal functions used to decode payload and sequentially store it into file stream.

- section with procedures containing encoded payload

Payload is represented as many (even thousands) of subsequent calls of method **sl**. Due to the procedure length limitation - calls are split into many individual procedures named **Pline0** .. **PlineN**

- function to be called to recreate and store payload

**saveFile** is the name of the function available for programmer, which save payload in provided *filename* path.

``` vbs
sub savefile( filename as String )
```

## Version history

- 0.1 [2018-12-15] Initial release

Embedded file encoded in base64 format.
