This script takes your manual MS Word Macro Code and applies it to any number of MS Word files using Python. Please follow the below checklist before starting the script:

1. No formatting is applied in Word files
2. There is no macro in Word file before running the script.
3. Check for '.docm' files. They should be in the folder
4. Macro text file location is correct. Location: "..\TXT Files\Macro_DirectCertifyQA.txt"

<u>Script Usage</u>

1. Code.py: Executes the script using ArgParse. It takes the `PATH` as the location of the script containing the `.docm` files

`python Code.py --path PATH`

2. CodeGUI.py: This script is used to generate the EXE file. It can also be executed in the console using:
`python CodeGUI.py`