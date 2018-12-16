# Payload2VBSConverter

## Description

Tool to embed payload (eg. malicious executable file) inside Microsoft office applications.

Presented tool allows for convertion of any type of the file into vbs script. Input file is encoded in base64 algorithm and represented as string constants. Generated output vbs code is embeddable into one of the MS office applications (Word, Excep, PowerPoint). The code contains functions to decode payload and store it back into file.

## Usage

The tool is written in Python version 2. Only standard build-in libraries are used, no additional modules must be installed. Run the tool by following command line:

```code
python payload2vbs_converter.py [filename]
```

### Arguments

- **filename** - path to file to be embedded inside generated output script.

## VBS limitations

Excel VBS limits size of individual string used as function parameters. If large amount of strings are presented inside the script, it Excel prevents from saving worksheet content. Therefore generated script uses relatively short string constants.
Another limitation is related to maximum number of lines in one procedure. The size of individual code blocks must be controlled. The tool automatically splits generated code into many separated procedures.

## Output code structure

