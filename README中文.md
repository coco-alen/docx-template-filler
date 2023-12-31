# word文档模板填充工具



### 1. 介绍

通过这个项目，可以设计一个模板.docx文件，使用{keywords}设置要填充的模板的位置，在另一个.xlsx文件中设置关键词内容，运行程序后自动完成填充。

比如，.docx文件模板被设定为：

```
{名字}今年{年龄}岁了，他是一个{性别}孩子
```

然后，设置一个 .xlsx表格填充关键词（注意，第一行作为提示不会被读取，备注列作为输入提示不会被读取）：

| 关键词 | 内容 | 备注               |
| ------ | ---- | ------------------ |
| 名字   | 张三 | 小孩的名字         |
| 年龄   | 5    | 小孩的年龄，纯数字 |
| 性别   | 男   | 男 / 女            |

将两个文件放到同一个路径下运行脚本，输入路径，输入.docx文件名（包含后缀），输入.xlsx文件名（包含后缀），程序会自动处理得到文本：

```
张三今年5岁了，他是一个男孩子
```





### 2 .docx 文件设置

你应该按照以下规则设置关键词：

```
这是一个测试文件，这是一个测试文件{关键词}，这是一个测试文件。
```

程序会尝试匹配{}中的关键词并替换为.xlsx中设置的内容。

- 若匹配到的关键词在xlsx中不存在，则不会处理，因此若需要使用大括号，仅需要保证大括号内不是有设置的关键词即可。
- 若出现大括号包裹的情况，如“这是{这是一个{关键词}测试文件}测试文件”，则仅会匹配最内部的{}内的文字作为关键词
- 不闭合的大括号将不进入匹配

![image-20230801020125538](./photo/image-20230801020125538.png)





### 3  .xlsx 文件设置

第一列设置关键字，第二列设置相应行中要替换的内容。

警告与提示：

- 第一行应用作输入提示，在程序运行时不会考虑。请勿修改第一行的提示内容
- 关键词列：用于匹配docx中被{}框定的关键词，如设定关键词“名字”，则会匹配{名字}，并将{名字}替换为内容列同一行的内容
- 备注列不会被程序读取，仅用于记忆关键词意义，建议注明要输入的内容格式，给出样例等以防止忘记
- 请勿设置完全相同的关键词，否则以最后一个设置为准，可以增加数字区分关键词
- 关键词排列可乱序，不必按照.docx中出现的顺序

![image-20230801020322165](./photo/image-20230801020322165.png)



### 4. 填充效果

![image-20230801020416387](./photo/image-20230801020416387.png)



### 5. 开发者设置

脚本使用正则化公式匹配关键词符号，若想使用不同的符号做关键词，可以在初始化TempFiller是设置自己的pattern参数。