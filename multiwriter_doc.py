from docxtpl import DocxTemplate
import os
from zhipuai import ZhipuAI
import streamlit as st
from docx import Document
from datetime import datetime, timedelta
import time
from io import BytesIO


class Multiwriter:

    def __init__(self,
                 api_key='edcaad198e032fdc4cbfdeabbf2662d4.ui4sERTRfZSUE6sy',
                 ) -> None:
        self.client = ZhipuAI(api_key=api_key)
        self.works_prompt = """你需要根据用户上传的工作职责生成当天的工作内容，字数严格控制在200字以内。"""
        self.ques_prompt = """你需要根据用户上传的工作内容生成当天工作中遇到的一个或者两个困难，字数严格控制在100字以内。"""
        self.method_prompt = """你需要根据用户上传的工作内容以及遇到的问题，生成一个解决思路，并阐述你取得的效果，字数控制在150字以内。"""
        self.parent_ptath = os.path.dirname(os.path.abspath(__file__))
    
    def generate_content(self,prompt,content):
        response = self.client.chat.completions.create(
            model="glm-4-airx",
            temperature=0.9,
            top_p=0.9,
            max_tokens=4095,
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": f"{content}"}
            ],
        )
        if response and response.choices and len(response.choices) > 0:
            return response.choices[0].message.content
        else:
            return 
    
    def generate_doc_content(self,job_name):
        work_content,ques_content,method_content = None,None,None
        times1, times2, times3 = 0,0,0
        while not work_content:
            if times1 <= 3:
                work_content = self.generate_content(self.works_prompt,job_name)
            else:
                break
        while not ques_content:
            if times2 <= 3:
                ques_content = self.generate_content(self.ques_prompt,work_content)
            else:
                break
        while not method_content:
            if times3 <= 3:
                method_content = self.generate_content(self.method_prompt,ques_content)
            else:
                break
        if work_content and ques_content and method_content:
            return work_content,ques_content,method_content
        else:
            return
        
    def write_content(self,date,time,weather,job_title,content,question_occur,method_effect,page,output_path):
        tpl = DocxTemplate(self.parent_ptath+os.sep+'实习工作日志正文.docx')
        context = {'date': date, 
        'time': time,
        'weather': weather,
        'job_title':job_title,
        'content':content,
        'question_occur':question_occur,
        'method_effect':method_effect}
        # 将标签内容填入模板中
        tpl.render(context)
        # 保存
        tpl.save(output_path+os.sep+f'{page}.docx')
    
    def remove_file(self,output_path,job_name,start_time,end_time):
        for file in os.listdir(output_path):
            if '.docx' in file:
                os.remove(output_path+os.sep+file)
                print(f'deleting {file}')
    
    def combine_doc(self,output_path,job_name,start_time,end_time,total_page):
        doc_path_list = []
        for file in range(total_page):
            if os.path.exits(output_path+os.sep+file+'.docx'):
                doc_path_list.append(output_path+os.sep+file+'.docx')
        merged_document = Document()
        for doc_name in doc_path_list:
            doc = Document(doc_name)
            doc.add_page_break()
            for element in doc.element.body:
                merged_document.element.body.append(element)
        buffer = BytesIO()
        merged_document.save(buffer)
        buffer.seek(0)
        merged_document.save(output_path+os.sep+f'{job_name} {start_time} - {end_time}.docx')
        return buffer

    
    def run(self):
        st.title('Library Worklog Generator')
        if True:
            job_title = st.text_input('职位名称 (Job Title)')
            start_time = st.date_input('开始时间 (Start Time)')
            end_time = st.date_input('结束时间 (End Time)')
            if st.button('提交'):
                results_placeholder = st.empty()
                if job_title and start_time and end_time:
                    #start_datetime_str = start_time
                    #end_datetime_str = end_time
                    #start_datetime = datetime.strptime(start_datetime_str, '%Y-%m-%d')
                    #end_datetime = datetime.strptime(end_datetime_str, '%Y-%m-%d')
                    datetime_list = []
                    current_datetime = start_time
                    while current_datetime <= end_time:
                        datetime_list.append(current_datetime)
                        current_datetime += timedelta(days=1)
                    total_page = len(datetime_list)
                    for _,date in enumerate(datetime_list):
                        results_placeholder.text(f'Generating {date} Workinglog Content....')
                        isoweekday_number = date.isoweekday()
                        isoweekday_name = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日'][isoweekday_number - 1]
                        work_content,ques_content,method_content = self.generate_doc_content(job_title+'实习生')
                        if os.path.exists(self.parent_ptath+os.sep+f'{job_title}'):
                            self.write_content(date,isoweekday_name,'晴',job_title,'\n'+work_content.strip(),'\n'+ques_content.strip(),'\n'+method_content.strip(),_,self.parent_ptath+os.sep+f'{job_title}')
                        else:
                            os.makedirs(self.parent_ptath+os.sep+f'{job_title}')
                            self.write_content(date,isoweekday_name,'晴',job_title,'\n'+work_content.strip(),'\n'+ques_content.strip(),'\n'+method_content.strip(),_,self.parent_ptath+os.sep+f'{job_title}')
                    results_placeholder.text(f'Combining Workinglog...')
                    docx_buffer = self.combine_doc(self.parent_ptath+os.sep+f'{job_title}',job_title,start_time,end_time,total_page)
                    self.remove_file(self.parent_ptath+os.sep+f'{job_title}',job_title,start_time,end_time)
                    results_placeholder.text(f'Finished!')
                    st.download_button(
                        label="Download DOCX",
                        data=docx_buffer.getvalue(),
                        file_name=f'{job_title} {start_time} - {end_time}.docx',
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    # 提示用户填写所有字段
                    st.error('请填写所有字段。')




if __name__ == '__main__':
    mw = Multiwriter()
    mw.run()
    #mw.combine_doc(mw.parent_ptath+os.sep+f'商业分析','商业分析',1,2)
    
