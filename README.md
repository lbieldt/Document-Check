# Document-Check
Tools to check documents

Current V1 features
- The program prompts the user to select a file (has to be docx)
- The program looks for the abbreviations that has double caps from [A-Z] and that are separated by "." and "&" i.e. B&W or A.I
- A new word document will be created at the user location (remember to save as docx)
- A table with all the abbreviations found with the occurrence number next to it will be generated.

To-do:
- Improve the prompt to only allow docx to be selected and also to save as docx
- Read all references that starts and ends with [] i.e. [1], [12] (save as a reference check)
- Create lookup table to fill in the expected abbreviation description
    -How to handle similar abbreviations i.e. SS - Security Systems or Stainless Steel?
- Check the document inline? So filling the Abbreviations table within the document (risky but will save a few steps)  