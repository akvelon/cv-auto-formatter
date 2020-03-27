# cv-auto-formatter
CV Auto Formatting tool. Reduces the daily toil for Client Services

## Approach:
The approach taken is to extract certain specified sections from a to-be-formatted-resume and then populate a static template. This template can be seen under ```template.docx```.

## To Use:

0. Run ```install_reqs.ps1``` to install required packages.

1. Place all to-be-formatted resumes in ```./INPUTS```.
2. Run ```run.ps1``` or the command ```python akvelon_format_enforcer.py```.
3. Formatted resumes will appear in the ```./RESULTS``` folder. Files for which there were issues will be place in ```./ISSUES``` unaltered.
4. Text logs are created and stored in ```./logs``` each time the code is ran.

## Notes

1. This current approach does not address the ```.doc``` file type. An initial start at addressing this file type was started and utilizes ```utils.py```. There is a helper function to automate the process of converting a ```.doc``` file to a ```.docx``` file.