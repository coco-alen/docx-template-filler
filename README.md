# docx-template-filler



### 1 - introduction

With this project, you can design a template .docx file, use {keywords} to set the position of the template to be filled, and set the keyword content in another .xlsx file, and automatically complete the filling after running the program

For example, .docx file template is set to:

```
{name} is {age} this year, and he is a {gender}
```

Then, set a .xlsx table to populate keywords (note that the first row will not be read as a hint, and the note column will not be read as an input hint):

| keyword | content | note       |
| ------- | ------- | ---------- |
| name    | Kevin   |            |
| age     | 13      |            |
| gender  | boy     | boy / girl |

Put two files under the same path to run the script, enter the path, enter the .docx file name (including the .docx), enter the .xlsx file name (including the .xlsx), and the program will automatically process the text:

```
Kevin is 13 this year, and he is a boy
```



### 2 - .docx file set

Your word template file should be set as follows：

```
this is a test file. this is a test file. {keyword} this is a test file. 
```

The program tries to match the {} character and read the key words in it.

- If the matched keyword does not exist in xlsx, it will not be processed, so if you need to use curly braces, you only need to make sure that there are no keywords set in the curly braces.
- In the case of curly braces, such as "This is {This is a {keyword} test file} test file", only the text in the innermost {} will be matched as a keyword
- Braces that do not close will not enter the match

![image-20230801020931261](.\photo\image-20230801020931261.png)



### 3 -  .xlsx file set

The first column sets the keywords, and the second column sets the content to be replaced in the corresponding row.

Warnings & Tips:

- The first line should be used as an input hint and will not be considered when the program is running. Do not modify the prompt content of the first line
- Keyword column: used to match keywords boxed by {} in docx, if the keyword "name" is set, {name} will be matched, and {name} will be replaced with the content of the same row as the content column
- The remarks column will not be read by the program, it is only used to remember the meaning of keywords, it is recommended to indicate the format of the content to be entered, give examples, etc. to prevent forgetting
- Do not set exactly the same keywords, otherwise the last setting will prevail, and you can add numbers to distinguish keywords
- Keywords can be arranged out of order, not necessarily in the order in which they appear in the .docx

![image-20230801021029683](.\photo\image-20230801021029683.png)

### 4. Result

![image-20230801021126153](.\photo\image-20230801021126153.png)



### 5. For developers

The script uses a regularization formula to match keyword symbols, if you want to use different symbols as keywords, you can set your own pattern parameters in the initialization TempFiller.