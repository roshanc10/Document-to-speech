# -*- coding: utf-8 -*-
"""
Created on Mon Feb 24 16:33:18 2020

@author: user
"""

def Text2Speech(text_result):
    language = 'en'
    audio_obj = gTTS(text = text_result, lang = language, slow = False)
    return audio_obj

import docxpy,os,sys,time
from gtts import gTTS
from glob import glob

def TextSpeechConverter(fldr_name):
    start_Time=time.time()
    for filename in glob(fldr_name+"*"):
        print(filename)
        text = docxpy.process(filename)

        path_of_output="C:/Users/user/Documents/"  #this is the path to generate file
        try:
            os.mkdir(os.path.join(path_of_output,"text"))
        except:
            pass

        text_file = filename.split("\\")[-1].replace(".docx",".txt")
        with open("C:/Users/user/Documents/text/"+text_file,"w",encoding='utf-8') as file:
            file.write(text) 
        try:
            os.mkdir(os.path.join(path_of_output,"audio"))
        except:
            pass

        audio_file = filename.split("\\")[-1].replace(".docx",".mp3")
        audio_obj = Text2Speech(text)
        audio_obj.save(os.path.join(path_of_output,"audio/")+audio_file)

        end_Time=time.time()
        execution_time=(end_Time-start_Time)/60
        print("Total Execution Time: ",execution_time)
        
#for using system arguments as folder name 
#if __name__ == '__main__':
#    fldr_name = sys.argv[1]
#    TextSpeechConverter(fldr_name)


#filename='D:/All_projects_data/Data Science/audio/dts/Input/Neural Networks and Deep Learning.docx'
fldr_name='C:/Users/user/Documents/input/'  #this is the input folder where all files to be processed should be kept.
TextSpeechConverter(fldr_name)

# extract text and write images in /tmp/img_dir
#text = docxpy.process(filename, "/tmp/img_dir"
                      
#if you want hyperlinks
#doc = docxpy.DOCReader(filename)
#doc.process()  # process file
#hyperlinks = doc.data['links']


