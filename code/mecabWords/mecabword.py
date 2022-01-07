import MeCab
import pandas as pd

def extract_words(file_path):
    file_in = open(file_path)
    f_line = file_in.read()
    mecab_tagger = MeCab.Tagger("-Ochasen")
    result=mecab_tagger.parse(f_line)
    my_list = []
    for i in result.splitlines()[:-1]:
        i = i.split()
        try:
            v = (i[2], i[1], i[-1])     
        except:
            pass
        my_list.append(v)

    word_dict = {}
    word_sub = {}
    word_pro={}
    for i in my_list:
        if i[-1].split('-')[0] not in ['助詞','記号']:
            if i[0] not in word_dict:
                word_dict[i[0]]=1
                word_sub[i[0]]=i[-1]
                word_pro[i[0]]=i[1]
            else:
                word_dict[i[0]] =word_dict[i[0]]+1
    df =pd.DataFrame({"fre":word_dict,'pro':word_pro,'sub':word_sub})
    df=df[df.fre>1]
    df=df.sort_values(by=['fre'],ascending=False)
    df.to_csv('extract_words.csv')

if __name__ == '__main__':
    file_path = input("输入文件读取地址: ")
    extract_words(file_path)