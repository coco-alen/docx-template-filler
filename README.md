# docx-template-filler



### 1.introduction

With this project, you can design a template .docx file, use {keywords} to set the position of the template to be filled, and set the keyword content in another .xlsx file, and automatically complete the filling after running the program



### 2. .docx file set

Your word template file should be set as follows：

```
this is a test file. this is a test file. {keyword} this is a test file. 
```

The program tries to match the {} character and read the key words in it.



### 3. .xlsx file set

The first column sets the keywords, and the second column sets the content to be replaced in the corresponding row.

Warning: The first line should be used as an input prompt and will not be considered when the program is running.