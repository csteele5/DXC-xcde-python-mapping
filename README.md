# XCDE JSON Map Creation

## Summary

This python utility program parses the contents of specific mapping configuration spreadsheets and creates usecase json maps for the XCDE configuration transport system.

## IDE

This repository is a complete virtual environment set up to run within PyCharm for development. All commits to GitHub are performed through PyCharm vs terminal.

## Prerequisites to run native

	- Install Python3 latest version [official website](https://www.python.org/downloads/)
		On Debian Linux - `sudo apt-get install python3`
	- Install ujson library using pip
		`python -m pip install ujson`
	- Install xlrd library using pip
		`python -m pip install xlrd`

	
To define what excel file to use open the file *createjson_gen.py* and change line 10.

To run the script, open the command prompt and execute "python3 createjson_gen.py" from the same directory as the createjson_gen.py file.