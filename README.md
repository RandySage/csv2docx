csv2docx
========

Python module for exporting a csv (selectively) to a docx, controlled by a json config file

License
-------
Available under the MIT License.  

I have no interest in restricting usage, so I gave this as much thought as clicking on 'I want it simple and permissive.'  Open an issue if you think this should be changed to another permissive license.  


Background
------
This module is intended to support other efforts.  I have a specification at work that is in Excel and I want to be able to export it to rich text.  

The concept is to encode column mappings and any other export content in a json, which this will ingest with the csv to produce the docx.  I'm sure this will evolve as time passes, but I'm likely to create a test folder that will necessarily show what happens, so I recommend you download and test it out.

### Tests
I will try to set up the tests folder so that you can set up a clean virtualenv, change to the tests directory, and run the following 
```
$VENV_PATH/bin/pip install -r requirements.txt
$VENV_PATH/bin/python $NAME_OF_TEST_FILE
```

python-docx vs pyrtf vs pyrtf-ng
-------
My initial exploration of pyrtf suggested that pyrtf-ng had all changes shown in pyrtf-ng as of Dec 23, 2013.  My initial attempt (~1 hour) to interactively repeat what I saw in the test cases in pyrtf-ng did not go smoothly, so I looked to see if there were other options.

I quickly came across python-docx (https://github.com/mikemaccana/python-docx), which had more contributors, more forks, and more recent contributors.  When I saw that there is pypy support (https://pypi.python.org/pypi/docx), 

I didn't give pyrtf any additional thought - I want to work directly in a web-app and I eventually want a variety of exports including docx, but in the meantime I'm stuck in a spreadsheet and my team wants to see docx anyway! 
