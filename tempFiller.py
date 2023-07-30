import re
import os

import docx
import openpyxl



class TempFiller():
    def __init__(self, dir, docx_fileName, excel_fileName, pattern=r"{(.*?)}"):

        if not docx_fileName.endswith(".docx"):
            raise Exception("word文件名必须以.docx结尾 (office word file name must END with .docx)")
        if not excel_fileName.endswith(".xlsx"):
            raise Exception("excel文件名必须以.xlsx结尾 (office excel file name must END with .xlsx)")

        self.input_file = os.path.join(dir, docx_fileName)
        self.output_file = os.path.join(dir, docx_fileName[:-5] + "-tempFiller.docx")
        self.keyword_file = os.path.join(dir, excel_fileName)

        if not os.path.exists(self.input_file):
            raise Exception("word源文件不存在 (word file not found)")
        if not os.path.exists(self.keyword_file):
            raise Exception("excel关键词文件不存在 (excel keyword file not found)")

        self.pattern = pattern

        try:
            self.docx = docx.Document(self.input_file)
            self.keyword_sheet = openpyxl.load_workbook(self.keyword_file).active
        except:
            raise Exception("文件打开失败 (File open failed)")
        
        print(f"导入模板文件：{self.input_file} (load template file)")
        print(f"导入关键词文件：{self.keyword_file} (load keyword file)")
        print(f"生成的文件输出到：{self.output_file} (output file to)")
        print("")

        self.keyword_dict = {}

    def load_keyword(self):
        # load keyword from excel file, and save to self.keyword_dict as dict
        for i, line in enumerate(self.keyword_sheet.rows):
            if line[0].value is not None and i != 0:
                self.keyword_dict[line[0].value] = [line[1].value, 0]  # set 0 to count

    def find_keyword(self, text):
        # find keyword in text, return a list
        match = re.findall(self.pattern, text)
        return match

    def relpace_text_iter(self, matched_keyword_list):
        # replace keyword in text, return a list
        for keyword in matched_keyword_list:
            if keyword in self.keyword_dict:
                self.keyword_dict[keyword][1] += 1
                yield [keyword, self.keyword_dict[keyword][0]]
            else:
                print(f"警告：关键词{keyword}不存在,跳过 (Warning: keyword {keyword} not found, skip)")
    
    def replace_in_runs(self, runs, matched_keyword_list):
        # have to replace in place, or the style will be lost
        # runs breaks up {} and the keywords in it, so need to try to identify them again
        keyword_record = ""
        keyword_begin = 0
        keyword_end = 0
        record_flag = False
        for i, run in enumerate(runs):
            if "{" in run.text:
                keyword_begin = i
                keyword_record = run.text
                record_flag = True
            else:
                if record_flag:
                    keyword_record += run.text
                    if "}" in run.text:
                        record_flag = False
                        keyword_end = i
                        keyword = self.find_keyword(keyword_record)[0]
                        
                        if keyword not in self.keyword_dict:
                            print(f"警告：关键词{keyword}不存在,跳过 (Warning: keyword {keyword} not found, skip)")
                        else:
                            self.keyword_dict[keyword][1] += 1
                            runs[keyword_begin].text = self.keyword_dict[keyword][0]
                            for j in range(keyword_begin + 1, keyword_end + 1):
                                runs[j].text = ""
            

    def replace_keyword(self):
        for para in self.docx.paragraphs:
            matched_keyword_list = self.find_keyword(para.text)
            if matched_keyword_list != []:
                self.replace_in_runs(para.runs, matched_keyword_list)

        for table in self.docx.tables:
            for row in table.rows:
                for cell in row.cells:
                    matched_keyword_list = self.replace_text(cell.text)
                    if matched_keyword_list != []:
                        for keyword, value in self.relpace_text_iter(matched_keyword_list):
                            cell.text = cell.text.replace("{" + keyword + "}", value)

    def save_docx(self):
        self.docx.save(self.output_file)

def main(temp_filler):
    temp_filler.load_keyword()
    temp_filler.replace_keyword()
    temp_filler.save_docx()

if __name__ == "__main__":
    dir = r"E:\Learning\tempFiller"
    input_file = r"测试文档.docx"
    keyword_file = r"测试表格.xlsx"

    temp_filler = TempFiller(dir=dir, docx_fileName=input_file, excel_fileName=keyword_file)
    main(temp_filler)

    # new_doc = replace_text_in_docx(input_file)
    # new_doc.save(output_file)