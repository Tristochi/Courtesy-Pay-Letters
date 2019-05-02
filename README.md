I am attempting to automate some manual processes at work in an attempt to learn more about Python and coding in general.
In this case I am taking a text file that is generated by a job we run in our financial platform. I cannot include the list because
it contains sensitive information but the format is like:

Column 1    Column 2    Column 3
Some info   Some numbers  Some info
etc         etc           etc
etc         etc           etc

Normally we would take this list, open it in excel, then separate the information into three other text files formatted the same way.
We then import those lists into a word document that has merge fields already on it. It is setup to use an sql statement to pull 
the info from the text files, but that does not always work. My aim in creating this is to make our lives easier, and something that
anyone else could implement at their job if they wanted. 

Currently my code is all over the place as I was trying to work out the mechanics first, and then refactor later. Also this is the first
independent project I have attempted in code outside of a tutorial. I am using the docx-mailmerge package which can be found here: 
https://pypi.org/project/docx-mailmerge/

My to-do list includes:
- Create a class for individuals
- Use properties stored in lists to assign them to objects
- Create a function to change info stored in the dictionaries, and add new ones if necessary.
(Dictionaries are required to merge info into the word doc, see https://pbpython.com/python-word-template.html)

I still have minimal idea of how to use git-hub so I will collect my thoughts and changes in the README I guess. 

-Tristochi